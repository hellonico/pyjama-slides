(ns ppt2md.core
  (:gen-class)
  (:import
    ;; PPTX (OOXML)
    (org.apache.poi.xslf.usermodel XMLSlideShow XSLFSlide XSLFShape XSLFTextShape XSLFTextParagraph XSLFTable XSLFTableCell)
    ;; PPT (binary)
    (org.apache.poi.hslf.usermodel HSLFTable HSLFTableCell HSLFSlideShow HSLFSlide HSLFShape HSLFTextShape)
    ;; Common
    (java.io File FileInputStream FileOutputStream BufferedWriter OutputStreamWriter)
    (java.nio.charset StandardCharsets)
    (java.util Locale)))

(defn- safe-trim [s]
  (some-> s str (.replace "\u000b" "\n") ; vertical tab → newline (seen in some decks)
          (.replace "\r" "\n")
          (.replaceAll "\n{3,}" "\n\n")
          (.trim)))

(defn- md-escape [s]
  ;; Escape characters that commonly affect Markdown structure.
  ;; Note: square brackets and parens must be escaped in the class.
  (-> s
      (clojure.string/replace #"([\\`*_\{\}\[\]\(\)#+\-!])" "\\\\$1")
      (clojure.string/replace #"\t" "    ")))


(defn- write-line! [sb s]
  (.append sb s) (.append sb "\n"))

(defn- bullet-line [level text]
  (let [indent (apply str (repeat level "  "))]
    (str indent "- " (md-escape (safe-trim text)))))

(defn- heading [level text]
  (let [lvl (max 1 (min 6 level))]
    (str (apply str (repeat lvl "#")) " " (md-escape (safe-trim text)))))

(defn- md-escape-table [s]
  ;; Extra escaping for table cells: escape `|` and flatten newlines.
  (-> (md-escape (safe-trim s))
      (clojure.string/replace #"\|" "\\\\|")
      (clojure.string/replace #"\n" "<br>")))

(defn- xslf-table->markdown [^XSLFTable tbl]
  (let [rows (.getRows tbl)
        rows-data (->> rows
                       (map (fn [row]
                              (->> (.getCells row)
                                   (map (fn [^XSLFTableCell c]
                                          (md-escape-table (.getText c))))
                                   (vec))))
                       (vec))
        maxcols (or (apply max 0 (map count rows-data)) 0)]
    (when (pos? maxcols)
      (let [pad (fn [r] (into [] (concat r (repeat (- maxcols (count r)) ""))))
            rows* (map pad rows-data)
            header (first rows*)
            body   (rest rows*)
            sep    (repeat maxcols "---")
            line   (fn [r] (str "| " (clojure.string/join " | " r) " |"))
            lines  (concat [(line header) (line sep)] (map line body))]
        (str "\n" (clojure.string/join "\n" lines) "\n")))))

(defn- hslf-table->markdown [^HSLFTable tbl]
  (let [rows (.getRows tbl)
        rows-data (->> rows
                       (map (fn [row]
                              (->> (.getCells row)
                                   (map (fn [^HSLFTableCell c]
                                          (md-escape-table (.getText c))))
                                   (vec))))
                       (vec))
        maxcols (or (apply max 0 (map count rows-data)) 0)]
    (when (pos? maxcols)
      (let [pad (fn [r] (into [] (concat r (repeat (- maxcols (count r)) ""))))
            rows* (map pad rows-data)
            header (first rows*)
            body   (rest rows*)
            sep    (repeat maxcols "---")
            line   (fn [r] (str "| " (clojure.string/join " | " r) " |"))
            lines  (concat [(line header) (line sep)] (map line body))]
        (str "\n" (clojure.string/join "\n" lines) "\n")))))



;; ---------- PPTX extraction ----------
(defn- pptx-paragraphs->md [^org.apache.poi.xslf.usermodel.XSLFTextParagraph p]
  (let [lvl (or (some-> p .getIndentLevel int) 0)
        text (-> p .getText safe-trim)]
    (when (seq text)
      (if (.isBullet p)
        (bullet-line lvl text)
        ;; Non-bulleted text becomes a paragraph (indented by level)
        (str (apply str (repeat lvl "  ")) (md-escape text))))))


(defn- pptx-shape->lines [^XSLFShape shape]
  (cond
    (instance? XSLFTextShape shape)
    (let [ts ^XSLFTextShape shape
          paras (.getTextParagraphs ts)]
      (->> paras (keep pptx-paragraphs->md)))


    (instance? org.apache.poi.xslf.usermodel.XSLFTable shape)
    (when-let [tbl-md (xslf-table->markdown ^org.apache.poi.xslf.usermodel.XSLFTable shape)]
      [tbl-md])  ;; return a single multi-line block

    :else
    ;; Skip pictures, charts, graphs, etc.
    nil))

(defn- read-pptx [^File f]
  (with-open [fis (FileInputStream. f)
              show (XMLSlideShow. fis)]
    (let [slides (.getSlides show)]
      (map-indexed
        (fn [idx ^XSLFSlide s]
          (let [sb (StringBuilder.)
                title (or (.getTitle s) (format "Slide %d" (inc idx)))
                shapes (.getShapes s)]
            (write-line! sb (heading 1 title))
            (doseq [sh shapes
                    :let [lines (pptx-shape->lines sh)]
                    :when (seq lines)]
              (doseq [ln lines] (write-line! sb ln)))
            (str sb)))
        slides))))

;; ---------- PPT (binary) extraction ----------
(defn- hslf-textshape-paragraphs [^HSLFTextShape ts]
  (->> (.getTextParagraphs ts)
       (mapcat identity) ; flattens List<List<HSLFTextParagraph>>
       (map (fn [p]
              (let [level (try (.getIndentLevel p) (catch Exception _ 0))
                    bullet? (try (.isBullet p) (catch Exception _ false))
                    text (safe-trim (try (.getText p) (catch Exception _ "")))]
                (when (seq text)
                  (if bullet?
                    (bullet-line (max 0 level) text)
                    (str (apply str (repeat (max 0 level) "  ")) (md-escape text)))))))))

(defn- hslf-shape->lines [^org.apache.poi.hslf.usermodel.HSLFShape shape]
  (cond
    (instance? org.apache.poi.hslf.usermodel.HSLFTextShape shape)
    (->> (hslf-textshape-paragraphs ^org.apache.poi.hslf.usermodel.HSLFTextShape shape)
         (remove nil?))

    (instance? org.apache.poi.hslf.usermodel.HSLFTable shape)
    (when-let [tbl-md (hslf-table->markdown ^org.apache.poi.hslf.usermodel.HSLFTable shape)]
      [tbl-md])

    :else
    nil))


(defn- read-ppt [^File f]
  (with-open [fis (FileInputStream. f)
              show (HSLFSlideShow. fis)]
    (let [slides (.getSlides show)]
      (map-indexed
        (fn [idx ^HSLFSlide s]
          (let [sb (StringBuilder.)
                title (or (.getTitle s) (format "Slide %d" (inc idx)))
                shapes (.getShapes s)]
            (write-line! sb (heading 1 title))
            (doseq [sh shapes
                    :let [lines (hslf-shape->lines sh)]
                    :when (seq lines)]
              (doseq [ln lines] (write-line! sb ln)))
            (str sb)))
        slides))))

;; ---------- Orchestrator ----------
(defn- ext [^String path]
  (-> path (.toLowerCase Locale/ROOT) (clojure.string/replace #"^.*\." "")))

(defn- to-md [^String in-path]
  (let [f (File. in-path)]
    (when-not (.exists f)
      (throw (ex-info "Input file not found" {:path in-path})))
    (case (ext in-path)
      "pptx" (clojure.string/join "\n" (read-pptx f))
      "ppt"  (clojure.string/join "\n" (read-ppt f))
      (throw (ex-info "Unsupported file type (use .ppt or .pptx)" {:path in-path})))))


(defn -main
  "Usage: clj -M -m ppt2md input.pptx [output.md]
   If output is omitted, prints Markdown to stdout."
  [& args]
  (try
    (when (< (count args) 1)
      (binding [*out* *err*]
        (println "Usage: clj -M -m ppt2md input.pptx [output.md]"))
      (System/exit 2))
    (let [[in-path out-path] args
          md (to-md in-path)]
      (if out-path
        (with-open [os (FileOutputStream. (File. out-path))
                    w (BufferedWriter. (OutputStreamWriter. os StandardCharsets/UTF_8))]
          (.write w md)
          (println "Wrote" out-path))
        (print md)))
    (catch Exception e
      (binding [*out* *err*]
        (println "Error:" (.getMessage e))
        (when-let [d (ex-data e)] (println d)))
      (System/exit 1))))
