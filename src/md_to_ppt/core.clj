(ns md-to-ppt.core
  (:require [clojure.java.io :as io]
            [clojure.string :as str]
            [ppt.core]
            [md.core]
            [pyjama.helpers.file :as hf])
  (:import (java.awt Rectangle)
           (java.io File FileOutputStream)
           (org.apache.poi.sl.usermodel PictureData$PictureType)
           (org.apache.poi.xslf.usermodel XMLSlideShow XSLFTextParagraph)))

(defn format-run [^XSLFTextParagraph para text]
  (let [tokens (-> text
                   (str/replace "**" "__B__")
                   (str/replace "*" "__I__")
                   (str/replace "`" "__C__")
                   (str/split #"(__[BIC]__)"))]
    (loop [segments tokens
           state {:bold false :italic false :code false}]
      (when (seq segments)
        (let [seg (first segments)]
          (cond
            (= seg "__B__") (recur (rest segments) (update state :bold not))
            (= seg "__I__") (recur (rest segments) (update state :italic not))
            (= seg "__C__") (recur (rest segments) (update state :code not))
            :else
            (let [run (.addNewTextRun para)]
              (.setText run seg)
              (.setFontSize run (if (:code state) 14.0 16.0))
              (.setFontFamily run (if (:code state) "Fira Code" "Inter"))
              (.setBold run (:bold state))
              (.setItalic run (:italic state))
              (recur (rest segments) state))))))))

(defn add-slide [ppt {:keys [type title body images tables code-blocks]}]
  (let [slide (.createSlide ppt)]
    ;; Title
    (when (some? title)
      (let [title-box (.createTextBox slide)
            para (.addNewTextParagraph title-box)
            run (.addNewTextRun para)]
        (.setAnchor title-box (Rectangle. 50 30 600 60))
        (.setText run (if (str/blank? title) " " title))
        (.setFontSize run 20.0)
        (.setBold run true)
        (.setFontFamily run "Inter")))

    ;; Body text
    (when (seq body)
      (let [box (.createTextBox slide)]
        (.setAnchor box (Rectangle. 50 100 600 200))
        (doseq [line body]
          (let [para (.addNewTextParagraph box)]
            (cond
              (re-find #"^\s*[-*] " line)
              (do
                (.setBullet para true)
                (let [txt (str/replace line #"^\s*[-*] " "")]
                  (format-run para txt)))

              :else
              (format-run para line))))))

    ;; Code blocks
    (doseq [{:keys [lang code]} code-blocks]
      (let [box (.createTextBox slide)]
        (.setAnchor box (Rectangle. 60 310 600 100))
        (let [para (.addNewTextParagraph box)]
          (.setBullet para false)
          (let [run (.addNewTextRun para)]
            (.setText run (str/join "\n" code))
            (.setFontFamily run "Fira Code")
            (.setFontSize run 12.0)
            (.setBold run false)
            (.setItalic run false)))))

    ;; Images
    (doseq [img-path images]
      (try
        (let [file (io/file img-path)
              data (.addPicture ppt (io/input-stream file) PictureData$PictureType/PNG)
              pic (.createPicture slide data)]
          (.setAnchor pic (Rectangle. 400 200 200 150)))
        (catch Exception e
          (println "⚠️  Failed to load image:" img-path))))
    (ppt.core/add-first-table-to-slide! slide tables {:header? true})))

(defn ^:private markdown-file? [^File f]
  (let [n (.getName f)]
    (or (.endsWith n ".md") (.endsWith n ".markdown"))))

(defn ^:private expand-patterns->files [patterns]
  (->> (hf/files-matching-path-patterns patterns)
       (filter markdown-file?)
       (sort-by #(.getCanonicalPath ^File %))))

(defn ^:private parse-cli-args [args]
  ;; supports: clj -M -m md-to-ppt.core [--out deck.pptx] <glob>...
  (let [out-idx (.indexOf ^java.util.List (vec args) "--out")]
    (if (neg? out-idx)
      {:out "combined.pptx" :patterns (seq args)}
      (let [out (get args (inc out-idx))
            pats (concat (take out-idx args) (drop (+ out-idx 2) args))]
        {:out      (or out "combined.pptx")
         :patterns (seq pats)}))))

(defn md->ppt!
  "Build one deck from an ordered seq of markdown file paths."
  [md-paths out-path]
  (let [ppt (XMLSlideShow.)]
    (doseq [md md-paths]
      (let [slides (md.core/parse-md md)]
        ;; optional: add a separator/title slide per file (uncomment if you want it)
        #_(add-slide ppt {:type :title :title (str (.getName (File. md)))})
        (doseq [s slides]
          (add-slide ppt s))))
    (with-open [out (FileOutputStream. out-path)]
      (.write ppt out))
    (println "✅ PowerPoint created at:" out-path)
    out-path))

(defn -main [& args]
  (let [{:keys [out patterns]} (parse-cli-args args)]
    (if (seq patterns)
      (let [files (expand-patterns->files patterns)
            paths (map #(.getPath ^File %) files)]
        (if (seq paths)
          (do
            (println "🧭 combining" (count paths) "markdown file(s) into" out)
            (md->ppt! paths out))
          (println "⚠️  No markdown files matched those patterns.")))
      (println "Usage:"
               "\n  clj -M -m md-to-ppt.core [--out deck.pptx] <path/glob> [<path/glob> ...]"
               "\nExamples:"
               "\n  clj -M -m md-to-ppt.core slides/**/*.md"
               "\n  clj -M -m md-to-ppt.core --out talk.pptx README.md notes/*.markdown"))))