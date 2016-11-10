(ns extra-large.core
  (:refer-clojure :exclude [read when-let get assoc! update!])
  (:require
   [better-cond.core :refer [when-let]]
   [camel-snake-kebab.core :refer [->kebab-case-keyword]]
   [dk.ative.docjure.spreadsheet :as doc]
   [extra-large.coords :as xl.coords]
   [extra-large.cell :as xl.cell]
   [extra-large.util :as util :refer [definstance?]]
   [clojure.set :as set]
   [clojure.string :as str]
   [clojure.spec.gen :as gen]
   [clojure.spec :as s]
   [clojure.spec.test :as test])
  (:import
   (org.apache.poi.ss.usermodel Workbook Sheet Cell Row)
   (org.apache.poi.xssf.usermodel XSSFWorkbook XSSFCell XSSFFormulaEvaluator)
   (org.apache.poi.ss.usermodel FormulaError)
   (org.apache.poi.ss.util CellRangeAddress)))

(definstance? poi-cell? Cell)

(definstance? poi-sheet? Sheet)

(definstance? poi-wb? Workbook)

(definstance? poi-row? Row)

(definstance? regex? java.util.regex.Pattern)

(s/def ::coords ::xl.coords/coords)

(s/def ::coords-range (s/tuple ::coords ::coords))

(s/def ::cell
  (s/keys :opt [::xl.cell/value
                ::xl.cell/formula]))

(def sheet-forbidden-chars (conj (set "[]/?*\\:") \u0000 \u0003))

(s/def ::sheet-name
  (s/and string? (partial every? (complement sheet-forbidden-chars))))

(def errors {:circular-ref             FormulaError/CIRCULAR_REF
             :div0                     FormulaError/DIV0
             :function-not-implemented FormulaError/FUNCTION_NOT_IMPLEMENTED
             :na                       FormulaError/NA
             :name                     FormulaError/NAME
             :null                     FormulaError/NULL
             :num                      FormulaError/NUM
             :ref                      FormulaError/REF
             :value                    FormulaError/VALUE})

(assert (= (s/form ::xl.cell/error)) (set (keys errors)))

(def cell-types {:blank   Cell/CELL_TYPE_BLANK
                 :num     Cell/CELL_TYPE_NUMERIC
                 :string  Cell/CELL_TYPE_STRING
                 :boolean Cell/CELL_TYPE_BOOLEAN
                 :error   Cell/CELL_TYPE_ERROR
                 :formula Cell/CELL_TYPE_FORMULA})

(defn new-wb [] (XSSFWorkbook.))

(s/fdef get-cell-type
  :args (s/cat :cell poi-cell?)
  :ret (set (keys cell-types)))

(defn get-cell-type [^Cell cell]
  ((set/map-invert cell-types) (.getCellType cell)))

(s/fdef set-cell-type!
  :args (s/cat :cell poi-cell? :type (set (keys cell-types)))
  :ret poi-cell?)

(defn set-cell-type! [^Cell cell type]
  (.setCellType cell (cell-types type)))

(s/fdef get-row
  :args (s/cat :sheet poi-sheet? :row ::xl.coords/row)
  :ret (s/nilable poi-row?))

(defn get-row
  [^Sheet poi-sheet row]
  (.getRow poi-sheet (dec row)))

(s/fdef get-or-create-row!
  :args (s/cat :sheet poi-sheet?
               :row ::xl.coords/row)
  :ret poi-row?)

(defn get-or-create-row!
  "Gets or creates the row in the sheet. Returns the poi-row.
  row is one based, like excel."
  [^Sheet poi-sheet row]
  (or (get-row poi-sheet row)
      (.createRow poi-sheet (dec row))))

(s/fdef get-cell
  :args (s/cat :row poi-row? :col ::xl.coords/col)
  :ret (s/nilable poi-cell?))

(defn ^Cell get-cell
  [^Row poi-row col]
  (.getCell poi-row (dec (xl.coords/col->int col))))

(s/fdef get-or-create-cell!
  :args (s/cat :row poi-row?, :col ::xl.coords/col)
  :ret poi-cell?)

(defn ^Cell get-or-create-cell!
  [^Row poi-row col]
  (or (get-cell poi-row col)
      (.createCell poi-row (dec (xl.coords/col->int col)))))

(defn read-wb
  "Reads a workbook from a stream or a file."
  [file-or-stream]
  (doc/load-workbook file-or-stream))

(defn write-wb!
  "Writes an excel workbook to a file or a stream.
  If it's a stream the caller must close the stream afterwards."
  [wb file-or-stream]
  (doc/save-workbook! file-or-stream wb))

(def docjure-errors
  #{:VALUE :DIV0 :CIRCULAR_REF :REF :NUM
    :NULL :FUNCTION_NOT_IMPLEMENTED :NAME :NA})

(s/def :xl.get-sheet/sheet
  (s/or :name string?
        :regex regex?
        :index nat-int?))

(s/def ::poi-args
  (s/alt :wb (s/cat :wb poi-wb?
                    :sheet (s/or :str string?
                                 :regex regex?
                                 :index nat-int?))
         :sheet poi-sheet?))

(s/fdef create-sheet
  :args (s/cat :wb poi-wb?, :name ::sheet-name)
  :ret poi-sheet?)

(defn create-sheet! [^Workbook poi-wb name]
  (.createSheet poi-wb name))

(s/fdef get-sheet
  :args ::poi-args
  :ret (s/nilable poi-sheet?))

(defmulti get-sheet
  (fn [_poi-wb sheet]
    (first (s/conform :xl.get-sheet/sheet sheet))))

(defmethod get-sheet :index
  [poi-wb sheet]
  (try (.getSheetAt ^Workbook poi-wb sheet)
       (catch IllegalArgumentException _ nil)))

(defmethod get-sheet :name
  [poi-wb sheet]
  (doc/select-sheet sheet poi-wb))

(defmethod get-sheet :regex
  [poi-wb sheet]
  (doc/select-sheet sheet poi-wb))

(s/fdef get-or-create-sheet!
  :args (s/cat :wb poi-wb? :sheet-name ::sheet-name)
  :ret poi-sheet?)

(defn get-or-create-sheet! [poi-wb sheet-name]
  (or (get-sheet poi-wb sheet-name)
      (create-sheet! poi-wb sheet-name)))

(s/fdef get-sheet-safe
  :args ::poi-args
  :ret poi-sheet?)

(defn get-sheet-safe
  "Will throw an IllegalArgumentexception if the sheet is not found."
  [poi-wb sheet]
  (or (get-sheet poi-wb sheet)
      (throw (IllegalArgumentException. "Sheet not found"))))

(s/fdef letsheets
  :args (s/cat :wb any?
               :sheet-names
               (s/and vector?
                      (partial apply distinct?)
                      (s/spec (s/+ (s/and simple-symbol?
                                          #(s/valid? ::sheet-name (name %))))))
               :body (s/* any?)))

(defmacro letsheets
  "Takes a workbook and a vector of sheet-names as symbols,
  binds the workbook to wb and binds the supplied symbols
  to the sheets they name (creates new sheets if no sheet found).

  (letsheets (xl/read-wb \"foo.xlsx\") [sales expenses profit]
    (xl/assoc! profit [:A 15] (- (xl/get-val sales [:C 59])
                                 (xl/get-val expenses [:C 45])))
    (xl/write-wb! wb \"bar.xlsx\")))"
  {:style/indent 2}
  [wb sheet-names & body]
  `(let [~'wb ~wb
         ~@(mapcat (fn [sheet-sym]
                     `[~sheet-sym (get-or-create-sheet! ~'wb ~(name sheet-sym))])
                   sheet-names)]
     ~@body))

(s/def ::cell-or-val
  (s/or :val ::xl.cell/value
        :cell ::cell))

(s/fdef force-formula-recalc!
  :args (s/cat :poi (s/or :wb poi-wb?
                          :sheet poi-sheet?)
               :recalculate? (s/? boolean?))
  :ret (s/or :wb poi-wb?
             :sheet poi-sheet?))

(defn force-formula-recalc!
  ([poi] (force-formula-recalc! poi true))
  ([poi recalculate?]
   (cond (poi-wb? poi) (.setForceFormulaRecalculation ^Workbook poi recalculate?)
         (poi-sheet? poi) (.setForceFormulaRecalculation ^Sheet poi recalculate?))
   poi))

(s/fdef get-poi
  :args (s/cat :poi ::poi-args
               :coords ::xl.coords/coords)
  :ret (s/nilable poi-cell?))

(defn get-poi
  "Get the poi cell at coords."
  ([poi-wb sheet coords]
   (get-poi (get-sheet-safe poi-wb sheet) coords))
  (^Cell [^Sheet poi-sheet [col row]]
   (when-let [poi-row (get-row poi-sheet row)]
     (get-cell poi-row col))))

(s/fdef get-poi!
  :args (s/cat :poi ::poi-args
               :coords ::xl.coords/coords)
  :ret poi-cell?)

(defn get-poi!
  "Get the poi cell at coords."
  ([poi-wb sheet coords]
   (get-poi (get-sheet-safe poi-wb sheet) coords))
  (^Cell [^Sheet poi-sheet [col row]]
   (-> poi-sheet
     (get-or-create-row! row)
     (get-or-create-cell! col))))

(s/fdef update-poi!
  :args (s/cat :poi ::poi-args
               :coords ::coords
               :f ifn?
               :args (s/* any?))
  :ret (s/or :wb poi-wb?
             :sheet poi-sheet?))

(defmulti update-poi! (fn [poi & _] (cond (poi-wb? poi) :wb
                                          (poi-sheet? poi) :sheet)))

(defmethod update-poi! :wb
  ^Workbook [poi-wb sheet & args]
  (apply update-poi! (get-sheet-safe poi-wb sheet) args)
  poi-wb)

(defmethod update-poi! :sheet
  [poi-sheet coords f & args]
  (apply f (get-poi! poi-sheet coords) args)
  poi-sheet)

(s/fdef get
  :args (s/cat :poi ::poi-args
               :coords ::xl.coords/coords)
  :ret ::cell)

(defn get
  "Get the cell at coords."
  ([wb sheet coords]
   (get (get-sheet-safe wb sheet) coords))
  ([^Sheet poi-sheet coords]
   (when-let [poi-cell (get-poi! poi-sheet coords)]

     (let [formula (try (.getCellFormula poi-cell)
                        ;; Throws exeption when the cell is not a formula cell
                        (catch java.lang.IllegalStateException _ nil))

           v (doc/read-cell poi-cell)

           v (if (keyword? v)
               (s/assert :xl.cell/error (->kebab-case-keyword v :separator "_"))
               v)]

       (-> #::xl.cell{:value v}
         (util/?> formula (assoc ::xl.cell/formula formula)))))))

(s/fdef get-val
  :args (s/cat :poi ::poi-args
               :coords ::xl.coords/coords)
  :ret ::xl.cell/value)

(defn get-val
  "Get the :xl.cell/value of the cell(s) at coords"
  ;; OPTIMIZE:
  ([poi-wb sheet coords]
   (::xl.cell/value (get poi-wb sheet coords)))
  ([poi-sheet coords]
   (::xl.cell/value (get poi-sheet coords))))

(s/fdef write-error-to-cell!
  :args (s/cat :cell poi-cell? :error ::xl.cell/error)
  :ret poi-cell?)

(defmulti coerce-cell-val (fn [value] (first (s/conform ::xl.cell/value value))))

(defmethod coerce-cell-val :blank   [_] nil)
(defmethod coerce-cell-val :num     [n] (double n))
(defmethod coerce-cell-val :string  [s] s)
(defmethod coerce-cell-val :boolean [b] b)
(defmethod coerce-cell-val :error   [e] (errors e))
;; TODO: What does excel do to dates.
(defmethod coerce-cell-val :date    [d] d)

(comment

  (def formula-cell? (comp boolean ::xl.cell/formula))

  (defn new-formula-evaluator ^XSSFFormulaEvaluator [^Cell poi-cell]
    (.. poi-cell getSheet getWorkbook
        getCreationHelper createFormulaEvaluator))

  (def wb (new-wb))

  (def foo (create-sheet! wb "foo"))

  ;; The formula seems to be evaluated when I setCellformula...
  (assoc! foo [:A 12] #::xl.cell{:formula "1 + 2"})

  )

(defn write-formula-and-val! [^Cell poi-cell value formula]
  (let [[value-type _] (s/conform ::xl.cell/value value)]
    (when formula
      (force-formula-recalc! (.. poi-cell getSheet getWorkbook)))

    (util/cond-doto poi-cell

      (not formula) (doc/set-cell! (coerce-cell-val value))

      ;; Is the formula evaluated automatically?
      formula (.setCellFormula formula))))

(defn write-cell!
  [^Sheet poi-sheet [col row] {::xl.cell/keys [value formula merged merged-by] :as cell}]
  (let [poi-cell (get-poi! poi-sheet [col row])
        [value-type v] (s/conform ::xl.cell/value value)]

    (write-formula-and-val! poi-cell value formula)

    poi-sheet))

(s/fdef assoc!
  :args (s/cat :poi ::poi-args
               :coords ::xl.coords/coords
               :cell ::cell-or-val)
  :ret (s/or :wb poi-wb? :sheet poi-sheet?))

(defn assoc!
  "Associates a value to a cell in the sheet."
  ([poi-wb sheet coords cell]
   (do (assoc! (get-sheet-safe poi-wb sheet) coords cell)
       poi-wb))
  ([^Sheet poi-sheet coords cell]
   (let [[value-type _] (s/conform ::cell-or-val cell)]
     (write-cell! poi-sheet coords (if (= :val value-type)
                                     #::xl.cell{:value cell} cell)))))

(s/fdef update-val!
  :args (s/cat :poi ::poi-args
               :coords ::xl.coords/coords
               :f ifn?
               :args (s/* any?))
  :ret (s/or :wb poi-wb?
             :sheet poi-sheet?))

(defn update-val!
  "Takes a a function that takes a ::xl.cell/value and
  replaces the cell with it's return value, which can be either a
  ::xl.cell/value or a ::xl/cell."
  {:arglists '([poi-wb sheet coords f & args]
               [poi-sheet coords f & args])}
  [& args]
  (let [{:keys [poi coords f] :as args*} (s/conform (-> `update! s/get-spec :args) args)
        coords ((juxt :col :row) coords)
        poi-sheet (case (first poi)
                    :sheet (second poi)
                    :wb (apply get-sheet-safe (take 2 args)))]
    (assoc! poi-sheet coords (apply f (get-val poi-sheet coords) (:args args*)))))

(s/fdef update!
  :args (s/cat :poi ::poi-args
               :coords ::xl.coords/coords
               :f ifn?
               :args (s/* any?))
  :ret (s/or :wb poi-wb?
             :sheet poi-sheet?))

(defn update!
  "Takes a a function that takes a ::xl/cell and
  replaces the cell with it's return value, which can be either a
  ::xl.cell/value or a ::xl/cell."
  {:arglists '([poi-wb sheet coords f & args]
               [poi-sheet coords f & args])}
  [& args]
  (let [{:keys [poi coords f] :as args*} (s/conform (-> `update! s/get-spec :args) args)
        coords ((juxt :col :row) coords)
        poi-sheet (case (first poi)
                    :sheet (second poi)
                    :wb (apply get-sheet-safe (take 2 args)))]
    (assoc! poi-sheet coords (apply f (get poi-sheet coords) (:args args*)))))

(s/fdef update-range!
  :args (s/cat :poi ::poi-args
               :range ::coords-range
               :f ifn?
               :args (s/* any?))
  :ret (s/or :wb poi-wb?
             :sheet poi-sheet?))

(defn update-range!
  "Updates a range of cells."
  {:arglists '([poi-wb sheet coords-range f & args]
               [poi-sheet coords-range f & args])}
  ([& args]
   (let [{:keys [poi f] :as args*} (s/conform (-> `update-range! s/get-spec :args) args)
         coords (mapv (juxt :col :row) (:range args*))
         poi-sheet (case (first poi)
                     :sheet (first args)
                     :wb (apply get-sheet-safe (take 2 args)))]
     (doseq [coord (apply xl.coords/range coords)]
       (apply update! poi-sheet coord f (:args args*)))
     (first args))))
