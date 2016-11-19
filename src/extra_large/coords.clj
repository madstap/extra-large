(ns extra-large.coords
  "Stuff related to columns and coordinates.
  A coordinate looks like [:AB 10] (equivalent to AB10 in excel)"
  (:refer-clojure :exclude [range])
  (:require
   [extra-large.util :as util :refer [fn->>]]
   [clojure.string :as str]
   [clojure.spec.test :as test]
   [clojure.spec.gen :as gen]
   [clojure.spec :as s]))

(declare col->int int->col)

(def alphabet "ABCDEFGHIJKLMNOPQRSTUVWXYZ")

;; The total number of available columns is 16K (2^14)
;; The total number of available rows is 1M (2^20)
(def max-cols (int (Math/pow 2 14)))
(def max-rows (int (Math/pow 2 20)))

(s/def ::col
  (s/with-gen (s/and simple-keyword?
                     #(every? (set alphabet) (name %))
                     #(<= (col->int %) max-cols))
              (fn []
                (gen/fmap int->col (s/gen (s/int-in 1 (inc max-cols)))))))

(s/def ::row (s/int-in 1 (inc max-rows)))

(s/def ::coords (s/tuple ::col ::row))

(s/def ::range (s/tuple ::coords ::coords))

(def ^:private alpha->int
  (zipmap alphabet (drop 1 (clojure.core/range))))

(def ^:private alpha-factors
  (iterate (partial *' 26) 1))

(s/fdef col->int
  :args (s/cat :col ::col)
  :ret nat-int?
  :fn (fn [{:keys [ret] {col :col} :args}]
        (= col (int->col ret))))

(s/fdef int->col
  :args (s/cat :n (s/int-in 1 (inc max-cols)))
  :ret ::col
  :fn (fn [{:keys [ret] {n :n} :args}]
        (= n (col->int ret))))

(defn col->int
  "Translates a column to an int. Starts at 1."
  [col]
  (->> (name col)
       reverse
       (map (fn [fac chr]
              (*' fac (alpha->int chr)))
            alpha-factors)
       (reduce +')))

(defn int->col
  "Translates a positive int to a column."
  [n]
  (loop [current n
         ret ()]
    (let [x (mod current 26)
          this (if (zero? x) 26 x)]
      (if (> current 0)
        (recur (bigint (/ (- current this) 26))
               (conj ret (char (+ 64 this))))
        (->> ret (apply str) keyword)))))

(s/def ::col-or-int (s/or :col ::col :int nat-int?))

(defn ->int [val]
  (let [[type val] (s/conform ::col-or-int val)]
    (case type :col (col->int val), :int val)))

(s/fdef col-inc
  :args (s/cat :col ::col)
  :ret ::col)

(defn col-inc
  "Increment a column"
  [col]
  (int->col (inc (col->int col))))

(s/fdef col-dec
  :args (s/cat :col (s/and ::col (complement #{:A})))
  :ret ::col)

(defn col-dec
  "Decrement a column.
  Cannot decrement :A as there are no negative columns."
  [col]
  (let [ret (dec (col->int col))]
    (assert (pos? ret))
    (int->col ret)))

(s/fdef col-min
  :args (s/+ ::col)
  :ret ::col)

(defn col-min
  "Returns the smallest of cols"
  [& cols]
  (int->col (apply min (map col->int cols))))

(s/fdef col-max
  :args (s/+ ::col)
  :ret ::col)

(defn col-max
  "Returns the largest of cols"
  [& cols]
  (int->col (apply max (map col->int cols))))

(s/fdef col+
  :args (s/+ ::col-or-int)
  :ret ::col)

(defn col+
  "Sums columns or ints together.
  Has no identity, so cannot take zero arguments."
  [& cols]
  (int->col (apply + (map ->int cols))))

(s/fdef col-
  :args (s/cat :init ::col-or-int
               :cols (s/+ ::col-or-int))
  :ret ::col)

(defn col-
  "Subtracts columns or ints.
  Needs at least two arguments, and will throw if the result is negative."
  [col1 col2 & cols]
  (let [ret (apply - (map ->int (cons col1 (cons col2 cols))))]
    (assert (pos? ret) "A column can't be negative.")
    (int->col ret)))

(s/fdef col>
  :args (s/+ ::col)
  :ret boolean?)

(defn col>
  "Like >, but with columns."
  [& cols]
  (apply > (map col->int cols)))

(s/fdef col<
  :args (s/+ ::col)
  :ret boolean?)

(defn col<
  "Like <, but with columns."
  [& cols]
  (apply < (map col->int cols)))

(s/fdef col<=
  :args (s/+ ::col)
  :ret boolean?)

(defn col<=
  "Like <=, but with columns."
  [& cols]
  (apply <= (map col->int cols)))

(s/fdef col>=
  :args (s/+ ::col)
  :ret boolean?)

(defn col>=
  "Like >=, but with columns."
  [& cols]
  (apply >= (map col->int cols)))

(s/fdef col-sort
  :args (s/cat :coll (s/coll-of ::col))
  :ret (s/and (s/coll-of ::col :kind seq?)
              (partial apply col<=)))

(defn col-sort [coll]
  (sort-by col->int coll))

(s/def :row-range/inc? boolean?)

(s/def :row-range/args
  (s/cat :args (s/? (s/alt :arity-1 (s/cat :end pos-int?)
                           :arity-2 (s/cat :start pos-int?
                                           :end pos-int?)
                           :arity-3 (s/cat :start pos-int?
                                           :end pos-int?
                                           :step pos-int?)))
         :opts (s/? (s/keys* :opt-un [:row-range/inc?]))))

(s/fdef row-range
  :args :row-range/args
  :ret (s/coll-of pos-int? :kind seq?))

(defn row-range
  "Takes the same arguments as range,
  but start, end and step are positive integers.
  Starts at 1. Also end inclusive by default.
  Takes optional keyword args.
  Options:

  :inc? Specifies whether the end should be inclusive.
        Start is always inclusive.
        Ignored if there's no end argument.
        Default: true"
  {:arglists '([& opts?]
               [end & opts?]
               [start end & opts?]
               [start end step & opts?])}
  [& args]
  (let [{[_ args*] :args
         {:keys [inc?] :or {inc? true}} :opts}
        (s/conform :row-range/args args)

        range-args ((util/juxtkeep :start :end :step)
                    (cond-> args*
                      (and inc? (:end args*)) (update :end inc)))]
    (->> range-args
      (apply clojure.core/range)
      (util/?>> (not (:start args*)) (drop 1)))))

(s/def :col-range/inc? boolean?)

(s/def :col-range/args
  (s/cat :args (s/? (s/alt :arity-1 (s/cat :end ::col)
                           :arity-2 (s/cat :start ::col
                                           :end ::col)
                           :arity-3 (s/cat :start ::col
                                           :end ::col
                                           :step pos-int?)))
         :opts (s/? (s/keys* :opt-un [:col-range/inc?]))))

(s/fdef col-range
  :args :col-range/args
  :ret (s/coll-of ::col :kind seq?))

(defn col-range
  "Takes the same arguments as range,
  but start and end are columns, step is a positive integer.
  Starts at :A. End inclusive by default.
  Takes optional keyword args.
  Options:

  :inc? Specifies whether the end should be inclusive.
        Start is always inclusive.
        Ignored if there's no end argument.
        Default: true"
  {:arglists '([& opts?]
               [end & opts?]
               [start end & opts?]
               [start end step & opts?])}
  [& args]
  (let [{[_ args*] :args
         {:keys [ret inc?] :or {inc? true}} :opts}
        (s/conform :col-range/args args)

        range-args ((util/juxtkeep :start :end :step)
                    (cond-> args*
                      (:start args*)          (update :start col->int)
                      (:end args*)            (update :end col->int)
                      (and (:end args*) inc?) (update :end inc)))]
    (->> range-args
      (apply clojure.core/range)
      (util/?>> (not (:start args*)) (drop 1))
      (map int->col))))

(s/def :range/by #{:row :col})

(s/fdef range
  :args (s/cat :coord1 ::coords
               :coord2 ::coords
               :by (s/? (s/keys* :opt-un [:range/by])))
  :ret (s/coll-of ::coords :kind seq?))

(defn range
  "Returns a sequence of the coords of the rectangle
  made using the points as corners.
  End inclusive, behaves like selecting stuff in excel.
  Starts in the upper left corner.
  Optional keyword arg :by specifies if its by column or by row, defaults to :row"
  [[c1 r1] [c2 r2] & {:keys [by] :or {by :row}}]
  (let [cols (col-range (col-min c1 c2) (col-max c1 c2))
        rows (row-range (min r1 r2)     (max r1 r2))]
    (case by
      :row (for [row rows, col cols] [col row])
      :col (for [col cols, row rows] [col row]))))

(def coords-regex #"[A-Z]+[0-9]+")

(def range-regex
  (->> coords-regex (repeat 2) (map str) (interpose ":") (apply str) re-pattern))

(s/def ::coords-str (s/and string? (partial re-find coords-regex)))

(s/def ::range-str (s/and string? (partial re-find range-regex)))

(s/fdef parse
  :args (s/cat :s ::coords-str)
  :ret ::coords)

(s/fdef parse-range
  :args (s/cat :s ::range-str)
  :ret ::range)

(s/fdef unparse-coords
  :args (s/cat :coords ::coords)
  :ret ::coords-str)

(s/fdef unparse-range
  :args (s/cat :range ::range)
  :ret ::range-str)

(defn parse [s]
  (let [[col row] (rest (re-find #"([A-Z]+)([0-9]+)" s))]
    [(keyword col) (Integer/parseInt row)]))

(defn parse-range [s]
  (mapv parse (str/split s #":")))

(defn unparse-coords [[col row]] (str (name col) row))

(defn unparse-range [range]
  (apply str (interpose ":" (map unparse-coords range))))
