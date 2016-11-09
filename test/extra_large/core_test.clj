(ns extra-large.core-test
  (:require [clojure.test :refer [deftest testing is are]]
            [clojure.spec :as s]
            [clojure.test.check.clojure-test :refer [defspec]]
            [clojure.spec.test :as test]
            [clojure.spec.gen :as gen]
            [clojure.test.check :as test.check]
            [clojure.test.check.properties :as prop]
            [extra-large.util :as util]
            [extra-large.core :as xl]
            [extra-large.cell :as xl.cell]))

(test/instrument)

(defspec can-write-and-read-back-values
  1000
  (prop/for-all [v (gen/such-that (complement inst?)
                     ;; TODO: How should dates work?
                     (s/gen ::xl.cell/non-error-value))
                 coords (s/gen ::xl/coords)]
    (xl/letsheets (xl/new-wb) [foo]
      (xl/assoc! foo coords v)
      (= (xl/coerce-cell-val v) (xl/get-val foo coords)))))
