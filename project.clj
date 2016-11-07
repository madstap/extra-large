(defproject extra-large "0.1.0-SNAPSHOT"
  :description "An excel library for clojure, wrapping apache poi"
  :url "http://github.com/madstap/extra-large"
  :license {:name "Eclipse Public License"
            :url "http://www.eclipse.org/legal/epl-v10.html"}
  :dependencies [[org.clojure/clojure "1.9.0-alpha14"]
                 [camel-snake-kebab "0.4.0"]
                 [better-cond "1.0.1"]
                 [org.apache.poi/poi "3.14"]
                 [org.apache.poi/poi-ooxml "3.14"]
                 [dk.ative/docjure "1.11.0"
                  :exclusions [org.apache.poi/poi
                               org.apache.poi/poi-ooxml]]]

  :profiles {:dev {:dependencies [[org.clojure/test.check "0.9.0"]]
                   :global-vars {*warn-on-reflection* true}}})
