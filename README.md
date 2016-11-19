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
;; Read a workbook from a file (or stream)
(def wb (xl/read "resources/foo.xlsx"))

;; Write a workbook to a file.
(xl/write! wb "resources/bar.xlsx")

;; Get a sheet
;; The sheet arg can either be a sheet name, regex or an index (0-based)
;; Will return nil when a sheet is not found,
;; there's also xl/get-sheet-safe which will throw.
(def sales (xl/get-sheet wb  #"(?i)sales"))

;; Create a sheet
(def expenses (xl/create-sheet! wb "Expenses"))

;; Get a sheet if it exists, if not create it. Only accepts a string.
(def foo (xl/get-sheet! wb "foo"))
```

#### Cells and coords

``` clojure
;; Cells are represented by a map.
#::xl.cell{:value 4.0 :formula "2 + 2"}

;; Coords by a tuple of column row
[:AB 42] ; The same as AB42 in excel, so one-based.
```

#### Doing stuff with cells

The main API consist functions that work with cells.

All of these functions have signatures that are like the following.
```clojure
;; Like xl/get-sheet, sheet-search can be a name, regex or a string.
;; Will throw if there's no sheet found.
([workbook sheet-search coords & maybe-more-args]

;; Just a poi sheet.
 [sheet coords & maybe-more-args])
```

Coords can be either coords `[:A 12]` or a range of coords `[[:A 1] [:B 5]]` (end inclusive).

`xl/assoc!`, `xl/update!`, `xl/update-val!` and `xl/update-poi!` are mutating functions
and will all return the workbook or sheet passed as the first argument,
as well as mutating it, so you can use them with ->. Calling them with a range
means that they'll be applied to each of the cells in the range in turn (by row).

`xl/get`, `xl/get-val`, `xl/get-poi` and `xl/get-poi!` are getters and will return (a) cell(s).
When given a range, there's a `:by` keyword argument that can be either `:row`(default) or `:col`.
The cells are returned in a two dimensional vector.

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

##### Calling the functions with ranges

``` clojure
;; To demonstrate I set the cell values to the string representation of their coords.
(doseq [coords (xl.coords/range [:A 1] [:B 2])]
  (xl/assoc! sales coords (xl.coords/unparse-coords coords)))

(xl/get-val sales [[:A 1] [:B 2]] :by :row) ;=> [["A1" "B1"] ["A2" "B2"]]

(xl/get-val sales [[:A 1] [:B 2]]) ;=> The same as above, default is by row.

(xl/get-val sales [[:A 1] [:B 2]] :by :col) ;=> [["A1" "A2"] ["B1" "B2"]]

```

## License

Copyright © 2016 Aleksander Madland Stapnes

Distributed under the Eclipse Public License either version 1.0 or (at
your option) any later version.
