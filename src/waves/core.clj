(ns waves.core
  (:require [clojure.java.io :as io]
            [clojure.string :as str]
            [pyjama.core])
  (:import (java.io FileInputStream FileOutputStream)
           (org.apache.poi.xslf.usermodel XMLSlideShow XSLFSlide XSLFTable XSLFTableCell XSLFTextShape)))


(defn translate [config text]
  (if (:debug config)
    (println ":> " text "\n"))

  (let [
        _options (if (contains? config :prompt)
                   config
                   (conj {:prompt (format (:prompt-template config) text)} config))
        ; TODO: pass through
        translation (clojure.string/trim
                      (pyjama.core/ollama
                        (:url config)
                        :generate
                        _options
                        :response
                        ))
        ]

    (if (:debug config)
      (println "< " translation "\n"))

    translation
    ))

(defn compute-output-file-path [input-file-path]
  (let [file (io/file input-file-path)
        parent-dir (.getParent file)
        original-name (.getName file)
        translated-name (str "translated_" original-name)]
    (str parent-dir "/" translated-name)))

(defn translate-many-ps [options paragraphs]

  (if (not (= :running (:status @options)))
    (throw (Exception. (str "Processing Interrupted:" @options))))

  (doseq [tp paragraphs]
    (let [paragraph-text (.getText tp)
          first-run (first (.getTextRuns tp))
          format {:font   (try (.getFontFamily first-run) (catch Exception e nil))
                  :size   (try (.getFontSize first-run) (catch Exception e nil))
                  ;:color (.getFillColor first-run)
                  :bold   (try (.isBold first-run) (catch Exception e nil))
                  :italic (try (.isItalic first-run) (catch Exception e nil))}
          translated-text (if (not (str/blank? paragraph-text)) (translate @options paragraph-text) "")
          ]
      (swap! options assoc-in [:processing]
             {:input paragraph-text, :output translated-text})

      ;; Clear the existing paragraph text
      (try
        ;; Use reflection to call the protected .clearButKeepProperties method
        (let [clear-method (.getDeclaredMethod (.getClass tp) "clearButKeepProperties" nil)]
          (.setAccessible clear-method true)
          (.invoke clear-method tp nil))

        (let [new-run (.addNewTextRun tp)]
          (.setText new-run translated-text)
          (when-let [font (:font format)] (.setFontFamily new-run font))
          (when-let [size (:size format)] (.setFontSize new-run size))
          ;(when-let [color (:color format)] (.setFillColor new-run color))
          (when (:bold format) (.setBold new-run true))
          (when (:italic format) (.setItalic new-run true)))

        (catch Exception e (.setText tp translated-text))))))


(defn translate-one-shape [options ^XSLFTextShape shape]
  (translate-many-ps options (.getTextParagraphs ^XSLFTextShape shape)))

(defn translate-one-cell [options ^XSLFTableCell cell]
  (let [tx-body (.getTextBody cell)]                        ;; Get the text body from the cell
    (when (some? tx-body)
      (translate-many-ps options (.getParagraphs tx-body)))))

(defn update-ppt-text [app-state]
  (let [input-path (:file-path @app-state)
        output-path (or (:output @app-state) (compute-output-file-path input-path))
        ;options @app-state
        ]
    (with-open [input-stream (FileInputStream. ^String input-path)
                output-stream (FileOutputStream. ^String output-path)]
      (let [ppt (XMLSlideShow. input-stream)]
        (doseq [^XSLFSlide slide (.getSlides ppt)]
          (doseq [shape (.getShapes slide)]
            (when (instance? XSLFTable shape)
              (doseq [row (.getRows ^XSLFTable shape)]
                (doseq [^XSLFTableCell cell (.getCells row)]
                  (translate-one-cell app-state cell))))
            (when (instance? XSLFTextShape shape)
              (translate-one-shape app-state shape))))
        (.write ppt output-stream)))))