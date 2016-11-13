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

(s/def ::coords-range ::xl.coords/coords-range)

(defn valid-formula? [{::xl.cell/keys [formula value]}]
  (let [[value-type _] (s/conform ::xl.cell/value value)]
    ;; Can't have an error in a non formula cell...
    (or (not= :error value-type) formula)))

(s/def ::cell
  (s/and valid-formula?
         (s/keys :opt [::xl.cell/value
                       ::xl.cell/formula
                       ::xl.cell/merged
                       ::xl.cell/merged-by])))

(def errors {:circular-ref             FormulaError/CIRCULAR_REF
             :div0                     FormulaError/DIV0
             :function-not-implemented FormulaError/FUNCTION_NOT_IMPLEMENTED
             :na                       FormulaError/NA
             :name                     FormulaError/NAME
             :null                     FormulaError/NULL
             :num                      FormulaError/NUM
             :ref                      FormulaError/REF
             :value                    FormulaError/VALUE})

(assert (= (s/form ::xl.cell/error) (set (keys errors))))

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

(s/fdef read
  :args (s/cat :file-or-stream any?)
  :ret poi-wb?)

(defn read
  "Reads a workbook from a stream or a file."
  [file-or-stream]
  (doc/load-workbook file-or-stream))

(defprotocol HasWorkbook
  (get-workbook [poi] "Gets the parent workbook of a poi object."))

(extend-protocol HasWorkbook
  Workbook
  (get-workbook [this] this)
  Sheet
  (get-workbook [^Sheet this] (.getWorkbook this))
  Cell
  (get-workbook [^Cell this] (.getWorkbook (.getSheet this))))

(defprotocol Writeable
  (write! [poi file-or-stream]
    "Writes a workbook or sheet to a file or a stream.
  If passed a sheet, will use the parent workbook.
  If passed a stream the caller must close the stream afterwards."))

(extend-protocol Writeable
  Workbook
  (write! [this file-or-stream] (doc/save-workbook! file-or-stream this))
  Sheet
  (write! [this file-or-stream] (write! (get-workbook this) file-or-stream)))

(def docjure-errors
  #{:VALUE :DIV0 :CIRCULAR_REF :REF :NUM
    :NULL :FUNCTION_NOT_IMPLEMENTED :NAME :NA})

(def sheet-forbidden-chars (conj (set "[]/?*\\:") \u0000 \u0003))

(s/def ::sheet-name
  (s/and string? (partial every? (complement sheet-forbidden-chars))))

(s/fdef create-sheet!
  :args (s/cat :wb poi-wb?, :name ::sheet-name)
  :ret poi-sheet?)

(defn create-sheet! [^Workbook poi-wb name]
  (.createSheet poi-wb name))

(s/def ::sheet-search
  (s/or :name ::sheet-name
        :regex regex?
        :index nat-int?))

(s/def ::get-sheet-args
  (s/cat :wb poi-wb?
         :sheet ::sheet-search))

(s/fdef get-sheet
  :args ::get-sheet-args
  :ret (s/nilable poi-sheet?))

(defmulti get-sheet
  "Finds a sheet in wb with either a sheet name, regex or sheet index.
  Returns nil is not found."
  {:arglists '([poi-wb sheet])}
  (fn [_ sheet]
    (first (s/conform ::sheet-search sheet))))

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

(s/fdef get-sheet!
  :args (s/cat :wb poi-wb? :sheet-name ::sheet-name)
  :ret poi-sheet?)

(defn get-sheet!
  "Finds a sheet in wb by sheet-name, creates the sheet if not found."
  [poi-wb sheet-name]
  (or (get-sheet poi-wb sheet-name)
      (create-sheet! poi-wb sheet-name)))

(s/fdef get-sheet-safe
  :args ::get-sheet-args
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

(s/fdef letsheets!
  :args (:args (s/get-spec `letsheets)))

(defn- sheet-bindings [get-sheet-fn sheet-names]
  (mapcat (fn [sheet-sym]
            `[~sheet-sym (~get-sheet-fn ~'wb ~(name sheet-sym))])
          sheet-names))

(defmacro letsheets
  "Takes a workbook and a vector of sheet-names as symbols,
  binds the workbook to wb and binds the supplied symbols
  to the sheets they name (throws an error if a sheet doesn't exist).

  (letsheets (xl/read-wb \"foo.xlsx\") [sales expenses profit]
    (xl/assoc! profit [:A 15] (- (apply + (xl/get-val sales [[:C 2] [:C 59]]))
                                 (apply + (xl/get-val expenses [[:C 2] [:C 45]]))))
    (xl/write! wb \"bar.xlsx\")))"
  {:style/indent 2}
  [wb sheet-names & body]
  `(let [~'wb ~wb
         ~@(sheet-bindings `get-sheet-safe sheet-names)]
     ~@body))

(defmacro letsheets!
  "Takes a workbook and a vector of sheet-names as symbols,
  binds the workbook to wb and binds the supplied symbols
  to the sheets they name (creates new sheets if no sheet found).

  (letsheets! (xl/read-wb \"foo.xlsx\") [sales expenses profit]
    (xl/assoc! profit [:A 15] (- (apply + (xl/get-val sales [[:C 2] [:C 59]]))
                                 (apply + (xl/get-val expenses [[:C 2] [:C 45]]))))
    (xl/write! wb \"bar.xlsx\")))"
  {:style/indent 2}
  [wb sheet-names & body]
  `(let [~'wb ~wb
         ~@(sheet-bindings `get-sheet! sheet-names)]
     ~@body))

(s/fdef cols-and-rows
  :args (s/cat :range ::coords-range)
  :ret (s/tuple (s/and (s/coll-of ::xl.coords/col)
                       (partial apply xl.coords/col<=))
                (s/and (s/coll-of ::xl.coords/row)
                       (partial apply <=))))

(defn cols-and-rows
  [coords-range]
  (let [[cols rows] (apply map vector coords-range)]
    [(xl.coords/col-sort cols) (sort rows)]))

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

(s/def ::poi-args
  (s/alt :wb ::get-sheet-args
         :sheet poi-sheet?))

(s/def ::coords-args
  (s/or :coords ::coords
        :range ::coords-range))

(s/def ::poi-and-coords-args
  (s/cat :poi ::poi-args
         :coords ::coords-args
         :more (s/* any?)))

(s/fdef cell-fn-dispatch
  :args ::poi-and-coords-args
  :ret #{:wb [:sheet :coords] [:sheet :range]})

(defn cell-fn-dispatch [& args]
  (let [{:keys [poi coords]} (s/conform ::poi-and-coords-args args)]
    (case (first poi)
      :wb :wb
      :sheet [:sheet (first coords)])))

(s/fdef get-poi
  :args (s/cat :poi ::poi-args
               :coords ::coords-args)
  :ret (s/nilable poi-cell?))

(defmulti get-poi
  "Get the poi cell at coords, or nil if no cell found.
  If called with a range, will return a lazy sequence of cells,
  with nils where no cells found."
  {:arglists '([poi-wb sheet coords cell-or-val]
               [poi-sheet coords cell-or-val])}
  cell-fn-dispatch)

(defmethod get-poi :wb
  [poi-wb sheet & args]
  (apply get-poi (get-sheet-safe poi-wb sheet) args))

(defmethod get-poi [:sheet :coords]
  ^Cell [poi-sheet [col row]]
  (when-let [poi-row (get-row poi-sheet row)]
    (get-cell poi-row col)))

(defmethod get-poi [:sheet :range]
  [poi-sheet coords-range]
  (let [[cols rows] (cols-and-rows coords-range)]

    (for [row (apply xl.coords/row-range rows)
          :let [poi-row (get-row poi-sheet row)]

          col (apply xl.coords/col-range cols)
          :let [poi-cell (and poi-row (get-cell poi-row col))]]

      poi-cell)))

(s/fdef get-poi!
  :args (s/cat :poi ::poi-args
               :coords ::coords-args)
  :ret poi-cell?)

(defmulti get-poi!
  "Get the poi cell at coords, will create the cell if no cell found.
  If called with a range, will return an _eager_ sequence of cells."
  {:arglists '([poi-wb sheet coords cell-or-val]
               [poi-sheet coords cell-or-val])}
  cell-fn-dispatch)

(defmethod get-poi! :wb
  [poi-wb sheet & args]
  (apply get-poi! (get-sheet-safe poi-wb sheet) args))

(defmethod get-poi! [:sheet :coords]
  ^Cell [poi-sheet [col row]]
  (-> poi-sheet
    (get-or-create-row! row)
    (get-or-create-cell! col)))

(defmethod get-poi! [:sheet :range]
  [poi-sheet coords-range]
  (let [[cols rows] (cols-and-rows coords-range)]
    ;; Should this be lazy?
    (doall
     (for [row (apply xl.coords/row-range rows)
           :let [poi-row (get-or-create-row! poi-sheet row)]

           col (apply xl.coords/col-range cols)
           :let [poi-cell (get-or-create-cell! poi-row col)]]
       poi-cell))))

(s/fdef update-poi!
  :args (s/cat :poi ::poi-args
               :coords ::coords-args
               :f ifn?
               :args (s/* any?))
  :ret (s/or :wb poi-wb?
             :sheet poi-sheet?))

(defmulti update-poi!
  "Runs the function with the poi cell at coords as the first arg.
  It presumably mutates the cell, the functions return value is ignored.
  When passed a range, maps the function over each cell.
  Returns the workbook or sheet passed as the first arg."
  {:arglists '([poi-wb sheet coords f & args]
               [poi-sheet coords f & args])}
  cell-fn-dispatch)

(defmethod update-poi! :wb
  [poi-wb sheet & args]
  (apply update-poi! (get-sheet-safe poi-wb sheet) args)
  poi-wb)

(defmethod update-poi! [:sheet :coords]
  ^Sheet [poi-sheet coords f & args]
  (apply f (get-poi! poi-sheet coords) args)
  poi-sheet)

(defmethod update-poi! [:sheet :range]
  ^Sheet [poi-sheet coords-range f & args]
  (run! #(apply f % args) (get-poi! poi-sheet coords-range))
  poi-sheet)

(s/fdef parse-merged-region
  :args (s/cat :s string?)
  :ret ::coords-range)

(defn parse-merged-region [s]
  (->> (str/split s #":")
    (map #(rest (re-find #"([A-Z]*)([[0-9]*])" %)))
    (mapv (juxt (comp keyword first) (comp util/str->int second)))))

(s/fdef sheet-merged-regions
  :args (s/cat :sheet poi-sheet?)
  :ret (s/coll-of ::coords-range :kind vector?))

(defn sheet-merged-regions [^Sheet poi-sheet]
  (mapv #(-> (.getMergedRegion poi-sheet %) .formatAsString parse-merged-region)
        (range (.getNumMergedRegions poi-sheet))))

(s/fdef find-merged-region
  :args (s/cat :sheet poi-sheet?
               :coords ::coords)
  :ret (s/nilable (s/spec (s/cat :idx nat-int? :region ::coords-range))))

(defn find-merged-region
  "Returns a tuple of [idx merged],
  where merged is the coords-range to which coords belongs,
  and idx is it's index in the merged-regions of the sheet.
  Returns nil if not found."
  [^Sheet poi-sheet coords]
  (reduce (fn [_ [idx merged-region]]
            (when (contains? (set (apply xl.coords/range merged-region)) coords)
              (reduced [idx merged-region])))
          nil (util/indexed (sheet-merged-regions poi-sheet))))

(defn ->cell-range ^CellRangeAddress [coords-range]
  (let [[cols rows] (apply map vector coords-range)
        cols (map xl.coords/col->int cols)
        [r1 r2 c1 c2] (map dec (concat rows cols))]
    (CellRangeAddress. (min r1 r2) (max r1 r2) (min c1 c2) (max c1 c2))))

(s/fdef get
  :args (s/cat :poi ::poi-args
               :coords ::coords-args)
  :ret ::cell
  :fn (fn [{:keys [args ret]}]
        (let [{::xl.cell/keys [merged merged-by]} ret
              {:keys [coords]} args]
          (cond merged (= coords (first merged))
                merged-by (contains? (set (rest merged-by)) coords)
                :else true))))

(defmulti get
  "Get the cell at coords. If passed a range, returns a lazy seq of cells."
  {:arglists '([poi-wb sheet coords]
               [poi-sheet coords])}
  cell-fn-dispatch)

(defmethod get :wb
  [poi-wb sheet & args]
  (apply get (get-sheet-safe poi-wb sheet) args))

(defmethod get [:sheet :coords]
  [^Sheet poi-sheet coords]
  (when-let [poi-cell (get-poi! poi-sheet coords)]
    (let [formula (try (.getCellFormula ^Cell poi-cell)
                       ;; Throws exeption when the cell is not a formula cell
                       (catch java.lang.IllegalStateException _ nil))

          [_ merged] (find-merged-region poi-sheet coords)

          merged-key (when merged
                       (if (= (first merged) coords)
                         ::xl.cell/merged
                         ::xl.cell/merged-by))

          v (doc/read-cell poi-cell)

          v (if (keyword? v)
              (s/assert :xl.cell/error (->kebab-case-keyword v :separator "_"))
              v)]

      (cond-> {}
        (not= ::xl.cell/merged-by merged-key) (assoc ::xl.cell/value v)
        formula (assoc ::xl.cell/formula formula)
        merged (assoc merged-key merged)))))

(defmethod get [:sheet :range]
  [poi-sheet coords-range]
  (map (partial get poi-sheet) (apply xl.coords/range coords-range)))

(s/fdef get-val
  :args (s/cat :poi ::poi-args
               :coords ::coords)
  :ret ::xl.cell/value)

(defmulti get-val
  "Get the value at coords.
  Returns a lazy sequence when provided a range."
  {:arglists '([poi-wb sheet coords]
               [poi-sheet coords])}
  cell-fn-dispatch)

(defmethod get-val :wb
  [poi-wb sheet & args]
  (apply get-val (get-sheet-safe poi-wb sheet) args))

(defmethod get-val [:sheet :coords]
  [poi-sheet coords]
  (::xl.cell/value (get poi-sheet coords)))

(defmethod get-val [:sheet :range]
  [poi-sheet coords-range]
  (map ::xl.cell/value (get poi-sheet coords-range)))

(defmulti coerce-cell-val
  "Coerces a value to the representation used in excel.
  Information can be lost in this process.
  For example: (coerce-cell-val 1/3) => 0.3333333333333333"
  (fn [value] (first (s/conform ::xl.cell/value value))))

(defmethod coerce-cell-val :blank   [_] nil)
(defmethod coerce-cell-val :num     [n] (double n))
(defmethod coerce-cell-val :string  [s] s)
(defmethod coerce-cell-val :boolean [b] b)
(defmethod coerce-cell-val :error   [e] (errors e))
;; TODO: What does excel do to dates.
(defmethod coerce-cell-val :date    [d] d)

(defn write-formula-and-val! [^Cell poi-cell value formula]
  (let [[value-type _] (s/conform ::xl.cell/value value)]
    (when formula
      (force-formula-recalc! (.. poi-cell getSheet getWorkbook)))

    (util/cond-doto poi-cell

      (not formula) (doc/set-cell! (coerce-cell-val value))

      ;; This will evaluate the formula
      formula (.setCellFormula formula))))

(defn remove-all-overlapping-regions!
  [^Sheet poi-sheet coord-range]
  (let [new-merged (apply xl.coords/range coord-range)]
    (reduce (fn [n-removed [idx merged-region]]
              (if (some (set (apply xl.coords/range merged-region)) new-merged)
                (do (.removeMergedRegion poi-sheet (- idx n-removed))
                    (inc n-removed))
                n-removed))
            0
            (util/indexed (sheet-merged-regions poi-sheet)))))

(defn write-merged! [^Sheet poi-sheet merged]
  (remove-all-overlapping-regions! poi-sheet merged)
  (.addMergedRegion poi-sheet (->cell-range merged)))

(defn write-cell!
  [^Sheet poi-sheet [col row] {::xl.cell/keys [value formula merged merged-by] :as cell}]
  (let [poi-cell (get-poi! poi-sheet [col row])
        [value-type v] (s/conform ::xl.cell/value value)
        mrgd (or merged merged-by)]

    (when mrgd (write-merged! poi-sheet mrgd))

    (write-formula-and-val! poi-cell value formula)

    poi-sheet))


(defn valid-merged-cell?
  [coords {::xl.cell/keys [merged merged-by]}]
  (cond merged (= coords (first merged))
        merged-by (contains? (set (rest (apply xl.coords/range merged-by)))
                             coords)
        :else true))

(s/fdef assoc!
  :args (s/and (s/cat :poi ::poi-args
                      :coords ::coords-args
                      :cell ::cell-or-val)
               (fn [{[coords-or-range coords] :coords
                     [cell-or-val cell]       :cell}]
                 (and (case coords-or-range
                        :coords true
                        :range
                        (not
                         ((some-fn ::xl.cell/merged ::xl.cell/merged-by) cell)))
                      (case cell-or-val
                        :val true
                        :cell (valid-merged-cell? coords cell)))))
  :ret (s/or :wb poi-wb? :sheet poi-sheet?))

(defmulti assoc!
  "Associates a value or cell to the sheet at coords."
  {:arglists '([poi-wb sheet coords cell-or-val]
               [poi-sheet coords cell-or-val])}
  cell-fn-dispatch)

(defmethod assoc! :wb
  [poi-wb sheet & args]
  (apply assoc! (get-sheet-safe poi-wb sheet) args)
  poi-wb)

(defmethod assoc! [:sheet :coords]
  [poi-sheet coords cell]
  (let [[value-type _] (s/conform ::cell-or-val cell)]
    (write-cell! poi-sheet coords (if (= :val value-type)
                                    #::xl.cell{:value cell} cell))
    poi-sheet))

(defmethod assoc! [:sheet :range]
  [poi-sheet coords-range cell]
  (doseq [coords (apply xl.coords/range coords-range)]
    (assoc! poi-sheet coords cell))
  poi-sheet)

(s/fdef update-val!
  :args (s/cat :poi ::poi-args
               :coords ::coords-args
               :f ifn?
               :args (s/* any?))
  :ret (s/or :wb poi-wb?
             :sheet poi-sheet?))

(defmulti update-val!
  "Takes a a function that takes a ::xl.cell/value and
  replaces the cell with it's return value, which can be either a
  ::xl.cell/value or a ::xl/cell."
  {:arglists '([poi-wb sheet coords f & args]
               [poi-sheet coords f & args])}
  cell-fn-dispatch)

(defmethod update-val! :wb
  [poi-wb sheet & args]
  (apply update-val! (get-sheet-safe poi-wb sheet) args)
  poi-wb)

(defmethod update-val! [:sheet :coords]
  [poi-sheet coords f & args]
  (assoc! poi-sheet coords (apply f (get-val poi-sheet coords) args))
  poi-sheet)

(defmethod update-val! [:sheet :range]
  [poi-sheet coords-range f & args]
  (run! (fn [coords]
          (apply update-val! poi-sheet coords f args))
        (apply xl.coords/range coords-range))
  poi-sheet)

(s/fdef update!
  :args (s/cat :poi ::poi-args
               :coords ::xl.coords/coords
               :f ifn?
               :args (s/* any?))
  :ret (s/or :wb poi-wb?
             :sheet poi-sheet?))

(defmulti update!
  "Takes a a function that takes a ::xl/cell and
  replaces the cell with it's return value, which can be either a
  ::xl.cell/value or a ::xl/cell."
  {:arglists '([poi-wb sheet coords f & args]
               [poi-sheet coords f & args])}
  cell-fn-dispatch)

(defmethod update! :wb
  [poi-wb sheet & args]
  (apply update! (get-sheet-safe poi-wb sheet) args)
  poi-wb)

(defmethod update! [:sheet :coords]
  [poi-sheet coords f & args]
  (assoc! poi-sheet coords (apply f (get poi-sheet coords) args))
  poi-sheet)

(defmethod update! [:sheet :range]
  [poi-sheet coords-range f & args]
  (run! (fn [coords]
          (apply update! poi-sheet coords f args))
        (apply xl.coords/range coords-range))
  poi-sheet)
