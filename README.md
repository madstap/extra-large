# extra-large

A nice clojure api wrapping apache poi,

## Status
Still figuring out the API. Pre-alpha.

I suggest using [docjure](https://github.com/mjul/docjure) for anything important.

## Usage

#### Deps

Requires clojure 1.9.0 which is currently `[org.clojure/clojure "1.9.0-alpha14"]`.

In your project.clj or build.boot

``` clojure
[extra-large "0.1.0-SNAPSHOT"]
```

In your namespace

``` clojure
(ns foo.core
  (:require
   [clojure.spec :as s]
   [extra-large.core :as xl]
   [extra-large.cell :as xl.cell]
   [extra-large.coords :as xl.coords]))
```

#### Use


Workbooks and sheets are the apache poi workbooks and sheets. They are mutable.

```clojure
;; Read from a file (or stream)
(def wb (xl/read-wb "resources/foo.xlsx"))

;; Write to a file.
(xl/write-wb! wb "resources/bar.xlsx")

;; Get a sheet
;; The sheet arg can either be a sheet name, regex or an index (0-based)
;; Will return zero when a sheet is not found,
;; there's also xl/get-sheet-safe which will throw.
(def sales (xl/get-sheet wb  #"(?i)sales"))

;; Create a sheet
(def expenses (xl/create-sheet! wb "Expenses"))
```

#### Cells and coords

``` clojure
;; Cells are represented by a map.
#::xl.cell{:value 4.0 :formula "2 + 2"}

;; Coords by a tuple of column row
[:AB 42] ; The same as AB42 in excel, so one-based.
```

#### Doing stuff with cells

All of these can take either the same two args as `xl/get-sheet` or a sheet as the first args.

`xl/assoc!`, `xl/update!`, `xl/update-val!` and `xl/update-poi!` will all return the workbook or sheet passed as the first argument.


##### Read cells

``` clojure
;; Get the value from the cell at A12
(xl/get-val sales [:A 12]) ;=> 42.0

;; Get the cell
(xl/get sales [:A 12]) ;=> #::xl.cell{:value 42.0 :formula "SUM(A1:A11)"}

```

##### Write stuff to cells
``` clojure
(xl/assoc! sales [:B 14] "bar")

(xl/get-val sales [:B 14]) ;;=> "bar"

;; Takes either a cell value or a cell
(xl/assoc! sales [:B 15] #::xl.cell{:formula "2 - 1"})

;; Formulas are evaluated
(xl/get sales [:B 15]) ;;=> #::xl.cell{:value 1.0 :formula "2 - 1"}
```

##### Update cells

``` clojure
;; update the value in a cell
(xl/update-val! sales [:B 14] (partial str "foo-"))

(xl/get sales [:B 14]) ;=> "foo-bar"

;; update a cell
(xl/update! sales [:A 12] (fn [{::xl.cell/keys [value formula] :as cell}]
                            (assoc cell ::xl.cell/formula "1 + 2")))

(xl/get sales [:A 12]) ;;=> #::xl.cell{:value 3 :formula "1 + 2"}

;; The return value from the function passed to xl/update-val! or xl/update!
;; can be either a cell or a value, like the last arg to xl/assoc!

```

There are also more low-level functions to work directly with apache poi cells.

``` clojure
;; get-poi will return either the cell at coords, or nil if there's no cell
(get-poi sales [:G 3]) ;=> nil

;; get-poi! will create the cell if not found.
(get-poi! sales [:G 3]) ;=> #object[org.apache.poi.xssf.usermodel.XSSFCell 0x552eef7e ""]

;; update-poi! will apply a function to the cell (will create the cell if it doesn't exist)
(update-poi! sales [:F 4] (fn [c] (.removeCellcomment c)))

```

## License

Copyright © 2016 Aleksander Madland Stapnes

Distributed under the Eclipse Public License either version 1.0 or (at
your option) any later version.
