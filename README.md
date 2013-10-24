# spreadmap

Evil Clojure library to turn Excel spreadsheets in
persistent reactive associative structures.

## Installation

```clj
[net.cgrand/spreadmap "0.1.4"]
```

### What's new in 0.1.4?
* The Java API with public modifiers on methods. *cough* *cough*

### What's new in 0.1.3?
* Due to popular demand: Java API! Call net.cgrand.SpreadMap.create(src) where src is a File, an InputStream or a String. It returns an Associative so you use .valAt and .assoc to read/update.
* Fix the value of FALSE cells (returned nil instead of false).

### What's new in 0.1.2?
* Rewrite to get rid of the crippled ForkedEvaluator from POI.
* It means XLSX are now supported as they should have been in the first place.

### What's new in 0.1.1?
* Keys can be "A1" "Foo!A1" "CellName" ["Foo" "A1"] [0 0] ["Foo" 0 0] [0 0 0]
* Dates are returned as j.u.Date instances

## Usage

```clj
=> (require '[net.cgrand.spreadmap :as evil])
nil
=> (def m (evil/spreadmap "/Users/christophe/Documents/Test.xls"))
#'user/m
=> (select-keys m ["A1" "B1" "C1"])
{"C1" 6.0, "B1" 2.0, "A1" 3.0}
; m is a spreadhseet where C1 is A1*B1
=> (-> m (assoc "B1" 12) (get "C1"))
36.0
=> (-> m (assoc "A1" 8) (get "C1"))
16.0
=> (defn mul [a b]
     (-> m
       (assoc "A1" a)
       (assoc "B1" b)
       (get "C1")))
#'user/mul
=> (mul 8 9)
72.0
=> (mul 7 6)
42.0
```

## License

Copyright © 2013 Christophe Grand

Distributed under the Eclipse Public License, the same as Clojure.
