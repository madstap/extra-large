(ns extra-large.coords-test
  (:require
   [extra-large.coords :as coords]
   [clojure.test.check.clojure-test :refer [defspec]]
   [extra-large.util :as util]
   [clojure.spec :as s]
   [clojure.spec.gen :as gen]
   [clojure.test.check.properties :as prop]
   [clojure.test.check :as test.check]
   [clojure.test :refer [deftest is are run-all-tests]]
   [clojure.spec.test :as test]))

(test/instrument)

(defspec int->col-roundtrip
  1000
  (prop/for-all [n (s/gen (s/int-in 1 (inc coords/max-cols)))]
    (= n (-> n coords/int->col coords/col->int))))

(defspec col->int-roundtrip
  1000
  (prop/for-all [c (s/gen ::coords/col)]
    (= c (-> c coords/col->int coords/int->col))))

(deftest cols-and-ints-test
  (are [col int] (and (= col (coords/int->col int))
                      (= int (coords/col->int col)))
    :A 1
    :Z 26
    :AA 27
    :BA 53))

(deftest col-inc-test
  (is (= :B (coords/col-inc :A)))
  (is (= :AA (coords/col-inc :Z))))

(deftest col-dec-test
  ;; The exact error that is thrown depends on whether col-dec is instrumented
  (is (thrown? #_AssertionError clojure.lang.ExceptionInfo (coords/col-dec :A)))
  (is (= :A (coords/col-dec :B))))

(deftest col-min-test
  (is (= :A (coords/col-min :C :B :A))))

(deftest col-max-test
  (is (= :C (coords/col-max :A :B :C))))

(deftest col+-test
  (is (= :F (coords/col+ :A 5)))
  (is (= :B (coords/col+ :A :A))))

(deftest col--test
  (is (= :A (coords/col- :B 1)))
  (is (= :A (coords/col- :C :B)))
  (is (thrown? AssertionError (coords/col- :B 2))))

(deftest col>-test
  (is (false? (coords/col> :B :A :A)))
  (is (true? (coords/col> :C :B :A)))
  (is (false? (coords/col> :A :B :C))))

(deftest col<-test
  (is (false? (coords/col< :A :A :B)))
  (is (true? (coords/col< :A :B :C)))
  (is (false? (coords/col< :C :B :A))))

(deftest col>=-test
  (is (true? (coords/col>= :A :A)))
  (is (false? (coords/col>= :A :B)))
  (is (true? (coords/col>= :B :A))))

(deftest col<=-test
  (is (true? (coords/col<= :A :A)))
  (is (true? (coords/col<= :A :B :B :C)))
  (is (false? (coords/col<= :C :B :A))))

(deftest row-range-test
  (is (= [1 2 3] (coords/row-range 3)))
  (is (= [1 2] (coords/row-range 3 :inc? false)))
  (is (= [10 11 12] (coords/row-range 10 12)))
  (is (= (range 10 22 2) (coords/row-range 10 20 2))))

(deftest col-range-test
  (is (= [:A :B :C] (coords/col-range :A :C) (coords/col-range :C)))
  (is (= [:A :B] (coords/col-range :A :C :inc? false)))
  (is (= [:A :D :G] (coords/col-range :A :G 3))))

(deftest range-test
  (is (= [[:A 1] [:B 1] [:A 2] [:B 2]] (coords/range [:A 1] [:B 2])))
  (is (= [[:A 1] [:A 2] [:B 1] [:B 2]] (coords/range [:A 1] [:B 2] :by :col)))
  (is (= [[:A 1]] (coords/range [:A 1] [:A 1]))))


(comment
  (run-all-tests)

  )
