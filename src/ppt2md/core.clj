(ns ppt2md.core
  (:gen-class)
  (:import
    ;; PPTX (OOXML)
    (org.apache.poi.xslf.usermodel XMLSlideShow XSLFSlide XSLFShape XSLFTextShape)
    ;; PPT (binary)
    (org.apache.poi.hslf.usermodel HSLFSlideShow HSLFSlide HSLFShape HSLFTextShape)
    ;; Common
    (java.io File FileInputStream FileOutputStream BufferedWriter OutputStreamWriter)
    (java.nio.charset StandardCharsets)
    (java.util Locale)))

;; ---------- Dynamic detection for optional HSLF table classes ----------
(def ^Class ?HSLFTable
  (try (Class/forName "org.apache.poi.hslf.usermodel.HSLFTable") (catch Throwable _ nil)))
(def ^Class ?HSLFTableCell
  (try (Class/forName "org.apache.poi.hslf.usermodel.HSLFTableCell") (catch Throwable _ nil)))

;; ---------- Utils ----------
(defn- safe-trim [s]
  (some-> s str
          (.replace "\u000b" "\n")  ;; vertical tab seen in some decks
          (.replace "\r" "\n")
          (.replaceAll "\n{3,}" "\n\n")
          (.trim)))

(defn- write-line! [^StringBuilder sb ^String s]
  (.append sb s) (.append sb "\n"))

(defn- heading [level text]
  (let [lvl (max 1 (min 6 level))]
    (str (apply str (repeat lvl "#")) " " (-> text safe-trim
                                              (clojure.string/replace #"([\\`*_\{\}\[\]\(\)#+\-!])" "\\\\$1")
                                              (clojure.string/replace #"\t" "    ")))))

(defn- bullet-line* [level preescaped]
  (let [indent (apply str (repeat level "  "))]
    (str indent "- " preescaped)))

(defn- indent [n s] (str (apply str (repeat n "  ")) s))

;; ---------- Escapers ----------
(defn- md-escape-general [s]
  ;; For non-table text: escape markdown specials.
  (-> (safe-trim s)
      (clojure.string/replace #"([\\`*_\{\}\[\]\(\)#+\-!])" "\\\\$1")
      (clojure.string/replace #"\t" "    ")))

(defn- md-escape-linktext [s]
  ;; For link TEXT: allow (), but escape brackets and common specials.
  (-> (safe-trim s)
      (clojure.string/replace #"([\\`*_{}#+\-!])" "\\\\$1")
      (clojure.string/replace #"\[" "\\\\[")
      (clojure.string/replace #"\]" "\\\\]")
      (clojure.string/replace #"\t" "    ")))

(defn- md-escape-table-inline [s]
  ;; For TABLE CELL runs: preserve **, _, +, -, etc. Only escape pipes and tabs.
  (-> (safe-trim s)
      (clojure.string/replace #"\|" "\\\\|")
      (clojure.string/replace #"\t" "    ")))

(defn- normalize-cell-newlines [s]
  (-> s (clojure.string/replace #"\n" "<br>")))

(defn- url->md-target [u]
  (let [u (safe-trim (str u))]
    (if (re-find #"[ )]" u) (str "<" u ">") u)))

;; ---------- Link-aware paragraph builders ----------
(defn- xslf-paragraph->text
  "Builds paragraph text with clickable links. Non-link runs are escaped with
   md-escape-general (regular text context)."
  [^org.apache.poi.xslf.usermodel.XSLFTextParagraph p]
  (let [runs (.getTextRuns p)]
    (->> runs
         (map (fn [^org.apache.poi.xslf.usermodel.XSLFTextRun r]
                (let [t (or (try (.getRawText r) (catch Exception _ nil))
                            (try (.getText r)     (catch Exception _ nil))
                            "")]
                  (if-let [hl (try (.getHyperlink r) (catch Exception _ nil))]
                    (when (seq (safe-trim t))
                      (when-let [addr (try (.getAddress hl) (catch Exception _ nil))]
                        (str "[" (md-escape-linktext t) "](" (url->md-target addr) ")")))
                    (md-escape-general t)))))
         (remove nil?)
         (apply str)
         safe-trim)))

(defn- xslf-paragraph->tabletext
  "Like xslf-paragraph->text but preserves inline markdown (**,_ etc.) for table cells."
  [^org.apache.poi.xslf.usermodel.XSLFTextParagraph p]
  (let [runs (.getTextRuns p)]
    (->> runs
         (map (fn [^org.apache.poi.xslf.usermodel.XSLFTextRun r]
                (let [t (or (try (.getRawText r) (catch Exception _ nil))
                            (try (.getText r)     (catch Exception _ nil))
                            "")]
                  (if-let [hl (try (.getHyperlink r) (catch Exception _ nil))]
                    (when (seq (safe-trim t))
                      (when-let [addr (try (.getAddress hl) (catch Exception _ nil))]
                        (str "[" (md-escape-linktext t) "](" (url->md-target addr) ")")))
                    (md-escape-table-inline t)))))
         (remove nil?)
         (apply str)
         safe-trim)))

(defn- hslf-paragraph->text [p] ;; p: HSLFTextParagraph
  (let [runs (try (.getTextRuns p) (catch Exception _ nil))]
    (->> runs
         (map (fn [r]
                (let [t (or (try (.getRawText r) (catch Exception _ nil))
                            (try (.getText r)     (catch Exception _ nil))
                            "")]
                  (if-let [hl (try (.getHyperlink r) (catch Exception _ nil))]
                    (when (seq (safe-trim t))
                      (when-let [addr (or (try (.getAddress hl) (catch Exception _ nil))
                                          (try (.getLabel   hl) (catch Exception _ nil)))]
                        (str "[" (md-escape-linktext t) "](" (url->md-target addr) ")")))
                    (md-escape-general t)))))
         (remove nil?)
         (apply str)
         safe-trim)))

(defn- hslf-paragraph->tabletext [p]
  (let [runs (try (.getTextRuns p) (catch Exception _ nil))]
    (->> runs
         (map (fn [r]
                (let [t (or (try (.getRawText r) (catch Exception _ nil))
                            (try (.getText r)     (catch Exception _ nil))
                            "")]
                  (if-let [hl (try (.getHyperlink r) (catch Exception _ nil))]
                    (when (seq (safe-trim t))
                      (when-let [addr (or (try (.getAddress hl) (catch Exception _ nil))
                                          (try (.getLabel   hl) (catch Exception _ nil)))]
                        (str "[" (md-escape-linktext t) "](" (url->md-target addr) ")")))
                    (md-escape-table-inline t)))))
         (remove nil?)
         (apply str)
         safe-trim)))

;; ---------- Tables ----------
(defn- rows->md-table [rows]
  (let [maxcols (or (apply max 0 (map count rows)) 0)]
    (when (pos? maxcols)
      (let [pad   (fn [r] (into [] (concat r (repeat (- maxcols (count r)) ""))))
            rows* (map pad rows)
            header (first rows*)
            body   (rest rows*)
            sep    (repeat maxcols "---")
            line   (fn [r] (str "| " (clojure.string/join " | " r) " |"))]
        (str "\n" (clojure.string/join "\n"
                                       (concat [(line header) (line sep)]
                                               (map line body)))
             "\n")))))

(defn- xslf-table->markdown [^org.apache.poi.xslf.usermodel.XSLFTable tbl]
  (let [rows (.getRows tbl)
        rows-data (->> rows
                       (map (fn [row]
                              (->> (.getCells row)
                                   (map (fn [^org.apache.poi.xslf.usermodel.XSLFTableCell c]
                                          (->> (.getTextParagraphs c)
                                               (map xslf-paragraph->tabletext)
                                               (remove clojure.string/blank?)
                                               (clojure.string/join "<br>")
                                               normalize-cell-newlines)))
                                   (vec))))
                       (vec))]
    (rows->md-table rows-data)))

(defn- hslf-table->markdown [tbl] ;; tbl: HSLFTable (detected dynamically)
  (let [rows (.getRows tbl)
        rows-data (->> rows
                       (map (fn [row]
                              (->> (.getCells row)
                                   (map (fn [cell]
                                          (let [paras (try (.getTextParagraphs cell) (catch Exception _ nil))]
                                            (normalize-cell-newlines
                                              (if paras
                                                (->> paras
                                                     (map hslf-paragraph->tabletext)
                                                     (remove clojure.string/blank?)
                                                     (clojure.string/join "<br>"))
                                                (md-escape-table-inline (str (try (.getText cell) (catch Exception _ ""))))))))
                                        vec)))
                            vec))]
    (rows->md-table rows-data)))

;; ---------- PPTX extraction ----------
(defn- pptx-paragraph->md [^org.apache.poi.xslf.usermodel.XSLFTextParagraph p]
  (let [lvl (or (some-> p .getIndentLevel int) 0)
        s   (xslf-paragraph->text p)]
    (when (seq s)
      (if (.isBullet p)
        (bullet-line* lvl s)      ;; s is already escaped per-run
        (indent lvl s)))))        ;; pre-escaped paragraph

(defn- pptx-shape->lines [^XSLFShape shape]
  (cond
    (instance? XSLFTextShape shape)
    (->> (.getTextParagraphs ^XSLFTextShape shape) (keep pptx-paragraph->md))

    (instance? org.apache.poi.xslf.usermodel.XSLFTable shape)
    (when-let [tbl-md (xslf-table->markdown shape)] [tbl-md])

    :else nil)) ;; skip pictures, charts, etc.

(defn- read-pptx [^File f]
  (with-open [fis (FileInputStream. f)
              show (XMLSlideShow. fis)]
    (map-indexed
      (fn [idx ^XSLFSlide s]
        (let [sb (StringBuilder.)
              title (or (.getTitle s) (format "Slide %d" (inc idx)))]
          (write-line! sb (heading 1 title))
          (doseq [sh (.getShapes s)
                  :let [lines (pptx-shape->lines sh)]
                  :when (seq lines)]
            (doseq [ln lines] (write-line! sb ln)))
          (str sb)))
      (.getSlides show))))

;; ---------- PPT (binary) extraction ----------
(defn- hslf-textshape-paragraphs [^org.apache.poi.hslf.usermodel.HSLFTextShape ts]
  (->> (.getTextParagraphs ts)
       (mapcat identity)
       (keep (fn [p]
               (let [lvl (try (int (.getIndentLevel p)) (catch Exception _ 0))
                     bul (try (.isBullet p) (catch Exception _ false))
                     s   (hslf-paragraph->text p)]
                 (when (seq s)
                   (if bul
                     (bullet-line* lvl s)
                     (indent lvl s))))))))

(defn- hslf-shape->lines [^org.apache.poi.hslf.usermodel.HSLFShape shape]
  (cond
    (instance? org.apache.poi.hslf.usermodel.HSLFTextShape shape)
    (->> (hslf-textshape-paragraphs ^org.apache.poi.hslf.usermodel.HSLFTextShape shape)
         (remove nil?))

    (and ?HSLFTable (.isInstance ?HSLFTable shape))
    (when-let [tbl-md (hslf-table->markdown shape)] [tbl-md])

    :else nil))

(defn- read-ppt [^File f]
  (with-open [fis (FileInputStream. f)
              show (HSLFSlideShow. fis)]
    (map-indexed
      (fn [idx ^HSLFSlide s]
        (let [sb (StringBuilder.)
              title (or (.getTitle s) (format "Slide %d" (inc idx)))]
          (write-line! sb (heading 1 title))
          (doseq [sh (.getShapes s)
                  :let [lines (hslf-shape->lines sh)]
                  :when (seq lines)]
            (doseq [ln lines] (write-line! sb ln)))
          (str sb)))
      (.getSlides show))))

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
  "Usage: clj -M -m ppt2md.core input.pptx [output.md]"
  [& args]
  (try
    (when (< (count args) 1)
      (binding [*out* *err*]
        (println "Usage: clj -M -m ppt2md.core input.pptx [output.md]"))
      (System/exit 2))
    (let [[in-path out-path] args
          md (to-md in-path)]
      (if out-path
        (with-open [os (FileOutputStream. (File. out-path))
                    w  (BufferedWriter. (OutputStreamWriter. os StandardCharsets/UTF_8))]
          (.write w md)
          (println "Wrote" out-path))
        (print md)))
    (catch Exception e
      (binding [*out* *err*]
        (println "Error:" (.getMessage e))
        (when-let [d (ex-data e)] (println d)))
      (System/exit 1))))
