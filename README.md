# spreadmap

Evil Clojure library to turn Excel spreadsheets in
persistent reactive associative structure.

## Installation

```clj
[net.cgrand/spreadmap "0.1.0"]
``

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
``

## License

Copyright © 2013 Christophe Grand

Distributed under the Eclipse Public License, the same as Clojure.
