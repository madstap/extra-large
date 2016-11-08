(ns extra-large.cell
  (:require
   [extra-large.util :as util]
   [clojure.spec.gen :as gen]
   [clojure.spec :as s]))

;; All the errors from
;; https://poi.apache.org/apidocs/org/apache/poi/ss/usermodel/FormulaError.html#enum_constant_summary
;; The same as the errors from docjure, just lower-kebab-cased.
(s/def ::error
  #{:num :function-not-implemented :ref
    :circular-ref :name :value :div0 :null :na})

;; Length of text cell contents is 32767
(def string-cell-max-length 32767)

(defn nan? [x]
  (and (number? x) (Double/isNaN x)))

(s/def ::number (s/and number?
                       (complement #{Double/NEGATIVE_INFINITY
                                     Double/POSITIVE_INFINITY})
                       (complement nan?)))

(s/def ::value (s/or :blank (s/with-gen (some-fn nil? #{""})
                                        #(gen/one-of [(gen/return nil)
                                                      (gen/return "")]))
                     :string (s/and string? #(<= (count %) string-cell-max-length))
                     :boolean boolean?
                     :num ::number
                     :error ::error
                     :date inst?))

(s/def ::non-error-value
  (s/and ::value (comp (partial not= :error) first)))

(s/def ::formula string?)
