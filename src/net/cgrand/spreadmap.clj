(ns net.cgrand.spreadmap
  (:require [clojure.java.io :as io])
  (:import [org.apache.poi.ss.usermodel Workbook WorkbookFactory CellValue DateUtil]
    [org.apache.poi.ss.formula.eval ValueEval StringEval BoolEval NumberEval BlankEval]
    org.apache.poi.ss.formula.eval.forked.ForkedEvaluator
    org.apache.poi.ss.formula.IStabilityClassifier
    [org.apache.poi.ss.util CellReference AreaReference]))

(defprotocol ValueEvalable
  (value-eval [v]))

(defprotocol Valueable
  (value [v wb cref]))

(defn- canon
  "Converts to a canonical cell ref [\"sheet\" row col]"
  [ref ^Workbook wb]
  (cond 
    (string? ref) 
    (recur
      (if-let [name (.getName wb ref)]
        (-> name .getRefersToFormula AreaReference. .getFirstCell)
        (CellReference. ^String ref))
      wb)
    (instance? CellReference ref)
    (let [^CellReference cref ref]
      [(or (.getSheetName cref) (-> wb (.getSheetAt 0) .getSheetName))
       (.getRow cref) (.getCol cref)])
    (= 3 (count ref))
    (let [[sheet row col] ref
          sheet (if (number? sheet) (-> wb (.getSheetAt sheet) .getSheetName) sheet)] 
      [sheet row col])
    (string? (second ref))
    (let [[sheet ^String ref] ref
          cref (CellReference. ref)
          sheet (if (number? sheet) (-> wb (.getSheetAt sheet) .getSheetName) sheet)] 
      [sheet (.getRow cref) (.getCol cref)])
    :else
    (let [[row col] ref
          sheet (-> wb (.getSheetAt 0) .getSheetName)] 
      [sheet row col])))

(defn- getter [^Workbook wb assocs]
  (let [evaluator
        (delay (let [evaluator
                     (ForkedEvaluator/create wb IStabilityClassifier/TOTALLY_IMMUTABLE nil)]
                 (doseq [[[sname row col] v] assocs]
                   (.updateCell evaluator sname row col v))
                 evaluator))]
    (fn [[sname row col :as cref]]
      (value (.evaluate ^ForkedEvaluator @evaluator sname row col)
        wb cref))))

(declare ss)

(deftype SpreadSheet [^Workbook wb assocs g]
  clojure.lang.Associative
  (assoc [this ref v]
    (ss wb (assoc assocs (canon ref wb) (value-eval v))))
  (containsKey [this ref]
    (boolean (.valAt this ref nil)))
  (entryAt [this ref]
    (when-let [v (.valAt this ref nil)]
      (clojure.lang.MapEntry. ref v)))
  clojure.lang.IPersistentCollection
  (cons [this x]
    (ss wb (into assocs
             (for [[ref v] (conj {} x)]
               [(canon ref wb) (value-eval v)]))))
  (equiv [this that]
    ; should be: same master and same assocs
    (.equals this that))
  clojure.lang.ILookup
  (valAt [this ref]
    (.valAt this ref nil))
  (valAt [this ref default]
    (or (-> ref (canon wb) g) default)))

(defn- ss [^Workbook wb assocs]
  (SpreadSheet. wb assocs (getter wb assocs)))

(defn spreadmap 
  "Creates a spreadmap from an Excel file, accepts same arguments as io/input-stream."
  [x & opts]
  (let [wb (with-open [^java.io.InputStream in
                       (apply io/input-stream x opts)]
             (WorkbookFactory/create in))]
    (ss wb {})))

(extend-protocol Valueable
  StringEval
  (value [v wb cref]
    (.getStringValue v))
  BoolEval
  (value [v wb cref]
    (.getBooleanValue v))
  NumberEval
  (value [v ^Workbook wb [sname row col]]
    (let [d (.getNumberValue v)]
      (if (some-> wb (.getSheet sname) (.getRow row) (.getCell col)
            DateUtil/isCellDateFormatted)
        (DateUtil/getJavaDate d)
        d)))
  BlankEval
  (value [v wb cref] nil))

(extend-protocol ValueEvalable
  ValueEval
  (value-eval [v] v)
  String
  (value-eval [v]
    (StringEval. v))
  Boolean
  (value-eval [v]
    (BoolEval/valueOf v))
  Number
  (value-eval [v]
    (NumberEval. (double v)))
  java.util.Date
  (value-eval [v]
    (NumberEval. (DateUtil/getExcelDate v)))
  nil
  (value-eval [v]
    BlankEval/instance))