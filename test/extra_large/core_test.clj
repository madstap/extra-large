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

(deftest get-workbook-test
  (xl/letsheets (xl/new-wb) [foo]
    (is (= (xl/get-workbook wb)
           (xl/get-workbook foo)
           (xl/get-workbook (xl/get-poi! foo [:A 1]))))))

(deftest merged-cells-overwrite-test
  (xl/letsheets (xl/new-wb) [foo]
    (xl/assoc! foo [:A 1] #::xl.cell{:merged [[:A 1] [:B 1]]})
    (xl/assoc! foo [:C 1] #::xl.cell{:merged [[:C 1] [:D 1]]})

    (testing "successfully merged cells"
      (is (= [[:A 1] [:B 1]] (::xl.cell/merged (xl/get foo [:A 1])))))

    (xl/assoc! foo [:B 1] #::xl.cell{:merged [[:B 1] [:C 1]]})

    (testing "Overwrites former merged regions"
      (is (= [[:B 1] [:C 1]] (::xl.cell/merged (xl/get foo [:B 1]))))
      (is (every? nil? (::xl.cell/merged [(xl/get foo [:A 1])
                                          (xl/get foo [:C 1])]))))))

(deftest merged-cells-validation-test
  (xl/letsheets (xl/new-wb) [foo]
    (testing "Can only merge the current cell"
      (is (thrown? clojure.lang.ExceptionInfo
                   (xl/assoc! foo [:A 1] #::xl.cell{:merged [[:B 1] [:C 1]]}))))

    (testing "::xl.cell/merged is only for the upper-left cell."
      (is (thrown? clojure.lang.ExceptionInfo
                   (xl/assoc! foo [:A 1] #::xl.cell{:merged-by [[:A 1] [:C 1]]})))

      (is (thrown? clojure.lang.ExceptionInfo
                   (xl/assoc! foo [:B 1] #::xl.cell{:merged [[:A 1] [:C 1]]}))))

    (testing "update! also validates merged regions"
      (is (thrown? clojure.lang.ExceptionInfo
                   (xl/update! foo [:A 1]
                               (constantly #::xl.cell{:merged [[:B 1] [:C 1]]})))))

    (testing "valid things are valid"
      (is (= [[:A 1] [:C 2]]
             (do (xl/assoc! foo [:A 1] #::xl.cell{:merged [[:A 1] [:C 2]]})
                 (::xl.cell/merged (xl/get foo [:A 1])))
             (::xl.cell/merged-by (xl/get foo [:B 1])))))))

(deftest a-formula-is-evaluated
  (xl/letsheets (xl/new-wb) [foo]
    (is (= #::xl.cell{:value 3.0 :formula "1 + 2"}
           (do (xl/assoc! foo [:A 1] #::xl.cell{:formula "1 + 2"})
               (xl/get foo [:A 1]))))))
