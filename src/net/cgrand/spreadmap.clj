(ns net.cgrand.spreadmap
  (:require [clojure.java.io :as io])
  (:import [org.apache.poi.ss.usermodel Workbook WorkbookFactory CellValue
            DateUtil Cell Row Sheet]
    [org.apache.poi.ss.formula.eval ValueEval StringEval BoolEval NumberEval BlankEval ErrorEval]
    [org.apache.poi.ss.formula IStabilityClassifier EvaluationWorkbook EvaluationSheet EvaluationName EvaluationCell FormulaParser FormulaType]
    [org.apache.poi.ss.util CellReference AreaReference]))

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

(defprotocol CellMisc
  (formula-tokens [cell wb]))

(defprotocol SheetMisc
  (sheet-index [sheet wb]))

(extend-protocol CellMisc
  EvaluationCell
  (formula-tokens [cell ^EvaluationWorkbook wb]
    (.getFormulaTokens wb cell)))

(defn- cell [^EvaluationSheet sheet idx row col v]
  (reify
    org.apache.poi.ss.formula.EvaluationCell
    (getSheet [this] sheet)
    (getCellType [this]
      (cond
        (instance? Boolean v) Cell/CELL_TYPE_BOOLEAN
        (number? v) Cell/CELL_TYPE_NUMERIC
        (string? v) Cell/CELL_TYPE_STRING
        (nil? v) Cell/CELL_TYPE_BLANK
        (:formula v) Cell/CELL_TYPE_FORMULA))
    (getNumericCellValue [this] (double v))
    (getIdentityKey [this] [idx row col])
    (getRowIndex [this] row)
    (getBooleanCellValue [this] v)
    #_(getErrorCellValue [this] )
    (getStringCellValue [this] v)
    (getColumnIndex [this] col)
    #_(getCachedFormulaResultType [this] (.getCachedFormulaResultType cell))
    CellMisc
    (formula-tokens [cell wb]
      (FormulaParser/parse (:formula v) wb FormulaType/CELL 
        idx))))

(extend-protocol SheetMisc
  EvaluationSheet
  (sheet-index [sheet ^EvaluationWorkbook wb]
    (.getSheetIndex wb sheet)))

(defn- sheet [^EvaluationSheet sht idx cells]
  (reify
     org.apache.poi.ss.formula.EvaluationSheet
     (getCell [this row col] 
       (if-let [kv (find cells [row col])]
         (cell this idx row col (val kv))
         (.getCell sht row col)))
     SheetMisc
     (sheet-index [this wb]
       (sheet-index sht wb))))

(defn- ^EvaluationWorkbook workbook [^EvaluationWorkbook wb assocs]
  (reify EvaluationWorkbook
    (getName [this G__3335 G__3336] (.getName wb G__3335 G__3336))
    (getName [this G__3337] (.getName wb G__3337))
    (getSheet [this idx] 
      (if-let [cells (some-> this (.getSheetName idx) assocs)]
        (sheet (.getSheet wb idx) idx cells)
        (.getSheet wb idx)))
    (getExternalName [this G__3339 G__3340] (.getExternalName wb G__3339 G__3340))
    (^int getSheetIndex [this ^EvaluationSheet sheet] (sheet-index sheet wb))
    (^int getSheetIndex [this ^String name] (.getSheetIndex wb name))
    (getSheetName [this G__3343] (.getSheetName wb G__3343))
    (resolveNameXText [this G__3344] (.resolveNameXText wb G__3344))
    (getUDFFinder [this] (.getUDFFinder wb))
    (convertFromExternSheetIndex [this G__3345] (.convertFromExternSheetIndex wb G__3345))
    (getExternalSheet [this G__3346] (.getExternalSheet wb G__3346))
    (getFormulaTokens [this cell] (formula-tokens cell wb))))

(defn- getter [wb assocs]
  (let [ewb (workbook
              (cond
               (instance? org.apache.poi.xssf.usermodel.XSSFWorkbook wb)
               (org.apache.poi.xssf.usermodel.XSSFEvaluationWorkbook/create wb)
               (instance? org.apache.poi.hssf.usermodel.HSSFWorkbook wb)
               (org.apache.poi.hssf.usermodel.HSSFEvaluationWorkbook/create wb))
              assocs)
        evaluator (org.apache.poi.ss.formula.WorkbookEvaluator.
                    ewb IStabilityClassifier/TOTALLY_IMMUTABLE nil)]
    (fn [[sname row col :as cref]]
      (when-let [cell (some-> ewb (.getSheet (.getSheetIndex ewb ^String sname))
                        (.getCell row col))]
        (value (.evaluate evaluator cell) wb cref)))))

(declare ss)

(deftype SpreadSheet [^Workbook wb assocs g]
  clojure.lang.Associative
  (assoc [this ref v]
    (let [[sname row col] (canon ref wb)] 
      (ss wb (assoc-in assocs [sname [row col]] v))))
  (containsKey [this ref]
    (boolean (.valAt this ref nil)))
  (entryAt [this ref]
    (when-let [v (.valAt this ref nil)]
      (clojure.lang.MapEntry. ref v)))
  clojure.lang.IPersistentCollection
  (cons [this x]
    (ss wb (into assocs
             (for [[ref v] (conj {} x)]
               [(canon ref wb) v]))))
  (equiv [this that]
    ; should be: same master and same assocs
    (.equals this that))
  clojure.lang.ILookup
  (valAt [this ref]
    (.valAt this ref nil))
  (valAt [this ref default]
    (or (@g (canon ref wb)) default)))

(defn- ss [^Workbook wb assocs]
  (SpreadSheet. wb assocs (delay (getter wb assocs))))

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

(defn fm= [formula-string] {:formula formula-string})

(defn- get-cells [ss sh row]
  (let [fc (.getFirstCellNum row)
        lc (.getLastCellNum row)
        rnum (.getRowNum row)]
    (map (fn [c] {(str (CellReference/convertNumToColString c))
                  (.valAt ss [sh rnum c])})
         (range fc lc))))

(defn read-sheet
  "Given a SpreadSheet and a sheet index, will read the entire sheet and
   return contents as a nested map. The outer map is keyed by the row number
   and the inner map by the column name."
  [^SpreadSheet ss ^java.lang.Integer idx]
  (let [sheet (-> (.wb ss) (.getSheetAt idx))
        fr (.getFirstRowNum sheet)
        lr (.getLastRowNum sheet)]
    (into {}
          (map (fn [r]
                 (let [row (.getRow sheet r)]
                   ;;inc row number as it is  0 based.
                   {(inc (.getRowNum row)) (into {} (get-cells ss idx row))}))
               (range fr (inc lr))))))
