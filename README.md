# extra-large

A nice clojure api wrapping apache poi,

## Status
Still figuring out the API. Pre-alpha.

I suggest using [docjure](https://github.com/mjul/docjure) for anything important.

## Usage

### Deps

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

### Use

#### Workbooks and sheets

Workbooks and sheets are apache poi workbooks and sheets. They are mutable.

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

;; Get a sheet if it exists, if not create it. Only accepts a string.
(def foo (xl/get-sheet! wb "foo"))

;; Create a sheet
(def expenses (xl/create-sheet! wb "Expenses"))
```

#### Cells and coords

``` clojure
;; Cells are represented by a map.
#::xl.cell{:value 4.0 :formula "2 + 2"}

;; Coords by a tuple of column row
[:AB 42] ; The same as AB42 in excel, so one-based.

;; A range of coords (end inclusive)
[[:A 1] [:C 10]]

```

#### Doing stuff with cells

The main API consist functions that work with cells.

Mutating functions
* `xl/assoc!`
* `xl/update!`
* `xl/update-val!`
* `xl/update-poi!`

Getters
* `xl/get`
* `xl/get-val`
* `xl/get-poi`
* `xl/get-poi!`

All of these functions have signatures that are like the following.
The first argument is always the poi object.
```clojure
;; Like xl/get-sheet, sheet-search can be a name, regex or a string.
;; Will throw if there's no sheet found.
([workbook sheet-search coords-or-range & maybe-more-args]

 [sheet coords-or-range & maybe-more-args])
```

The mutators all return the workbook or sheet passed as the first argument,
as well as mutating it, so you can use them with `->`. Calling them with a range
means that they'll be applied to each of the cells in the range in turn (by row).

The getters return a cell (value, cell-map or poi-cell) when called with coords,
and a vector of vectors of cells when called with a range.
The `:by` keyword argument controls whether it's a vector of rows or columns.
`:by :row` or `:by :col`, default `:row`.

#### Read cells

``` clojure
;; Get the value from the cell at A12
(xl/get-val sales [:A 12]) ;=> 42.0

;; Get the cell
(xl/get sales [:A 12]) ;=> #::xl.cell{:value 42.0 :formula "SUM(A1:A11)"}

```

#### Write stuff to cells
``` clojure
(xl/assoc! sales [:B 14] "bar")

(xl/get-val sales [:B 14]) ;;=> "bar"

;; Takes either a cell value or a cell
(xl/assoc! sales [:B 15] #::xl.cell{:formula "2 - 1"})

;; Formulas are evaluated
(xl/get sales [:B 15]) ;;=> #::xl.cell{:value 1.0 :formula "2 - 1"}
```

#### Update cells

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
(update-poi! sales [:F 4] #(.removeCellcomment %))
```

#### Calling getters with a range

``` clojure
(doseq [coords (xl.coords/range [:A 1] [:B 2])]
  (xl/assoc! sales coords (xl.coords/unparse coords)))

;; Default is by row
(xl/get-val sales [[:A 1] [:B 2]]) ;=>          [["A1" "B1"]
                                   ;             ["A2" "B2"]]

(xl/get-val sales [[:A 1] [:B 2]] :by :col) ;=> [["A1" "A2"]
                                            ;    ["B1" "B2"]]

```

## License

Copyright © 2016 Aleksander Madland Stapnes

Distributed under the Eclipse Public License either version 1.0 or (at
your option) any later version.
