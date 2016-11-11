(ns extra-large.util
  (:require
   [clojure.spec :as s]
   [clojure.spec.test :as test]))

(defn pp< [x] (clojure.pprint/pprint x) x)

(defn ppf< [f x] (clojure.pprint/pprint (f x)) x)

(defmacro fn->
  "Same as #(-> % ~@forms)"
  [& forms]
  `(fn [x#] (-> x# ~@forms)))

(defmacro fn->>
  "Same as #(->> % ~@forms)"
  [& forms]
  `(fn [x#] (->> x# ~@forms)))

(defmacro ?>
  "Conditional single-arrow operation (-> m (?> add-kv? (assoc :k :v)))"
  [x test & forms]
  `(if ~test
     (-> ~x ~@forms)
     ~x))

(defmacro ?>>
  "Conditional double-arrow operation (->> nums (?>> inc-all? (map inc)))"
  [test & forms]
  `(if ~test
     (->> ~(last forms) ~@(butlast forms))
     ~(last forms)))

(s/fdef juxtkeep
  :args (s/+ ifn?)
  :ret (s/fspec :ret (s/and vector? (partial every? some?))))

(defn juxtkeep
  "Like juxt, but removes any nil values from the result."
  [& fns]
  (fn [& args]
    (vec (keep #(apply % args) fns))))

(s/fdef str->int
  :args (s/cat :x (s/or :str string?
                        :chr char?))
  :ret int?)

(defn str->int [x]
  (Integer/parseInt (str x)))

(s/fdef cond-doto
  :args (s/cat :x any?
               :clauses (s/* (s/cat :test any? :expr any?))))

(defn indexed [coll]
  (map vector (range) coll))

(defmacro cond-doto
  "Takes a presumably mutable x and zero or more clauses.
  A clause is a test expr, which will be evaluated and checked for truthyness,
  and an expression which "
  {:style/indent 1}
  ([x] x)
  ([x & clauses]
   (let [[test expr & more] clauses]
     `(cond-doto (if ~test (doto ~x ~expr) ~x) ~@more))))

(s/fdef definstance?
  :args (s/cat :name-sym simple-symbol?
               :docstring (s/? string?)
               :meta (s/? map?)
               :class any?))

(defmacro definstance?
  {:style/indent 1}
  ([& args]
   (let [{:keys [name-sym docstring meta class]}
         (s/conform (:args (s/get-spec `definstance?)) args)]
     `(do
        (s/fdef ~name-sym :args (s/cat :x any?) :ret boolean?)
        (defn ~name-sym ~@(cond-> []
                            docstring (conj docstring)
                            meta (conj meta)
                            true (into [['x] `(instance? ~class ~'x)])))))))

(comment

  (test/instrument)

  )
