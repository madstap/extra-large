(ns extra-large.util-test
  (:require
   [extra-large.util :refer :all]
   [clojure.test :refer [deftest is are]]
   [clojure.spec :as s]))

(deftest fn->-test
  (is (= {:a 11, :b 4}
         ((fn-> (update :a inc) (assoc :b 5) (update :b dec)) {:a 10}))))

(deftest fn->>-test
  (is (= (range 0 12 2)
         ((fn->> (take 11) (remove odd?)) (range 20)))))

(deftest ?>>-test
  (is (= (range 0 88 8)
         (->> (range 40)
           (take 22)
           (?>> true (filter even?) (map (partial * 2)))
           (?>> (zero? 123) (map (partial * -1)))
           (map (partial * 2))))))

(deftest juxtkeep-test
  (is (= [0 1] ((juxtkeep :a :b :c) {:a 0 :b nil :c 1}))))

(deftest cond-doto-test
  (= [1 3] @(cond-doto (atom [])
              true (swap! conj 1)
              false (swap! conj 2)
              true (swap! conj 3)))

  (is (let [f #(swap! % conj 1)]
        (= [1] @(cond-doto (atom [])
                  true f))))

  (= :a @(cond-doto (atom :a))))

(comment

  (clojure.test/run-all-tests)

  )
