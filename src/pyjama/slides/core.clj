(ns pyjama.slides.core
  (:require [cheshire.core :as json]
            [clojure.java.io :as io]
            [clojure.string :as str]
            [pyjama.core :as pyjama])
  (:import (java.io FileOutputStream)
           (org.apache.poi.xslf.usermodel XMLSlideShow)))

(defn parse-outline
  "Parse a text outline into a structured format for slides"
  [outline]
  (let [lines (str/split-lines outline)
        slides (atom [])
        current-slide (atom nil)]

    (doseq [line lines]
      (cond
        ;; New slide title (assuming titles start with # or are all caps)
        (or (str/starts-with? line "#")
            (and (not (str/blank? line))
                 (= (str/upper-case line) line)))
        (do
          (when @current-slide
            (swap! slides conj @current-slide))
          (reset! current-slide {:title  (str/replace line #"^#\s*" "")
                                 :points []}))

        ;; Bullet point (assuming they start with - or *)
        (re-find #"^\s*[-*]\s+" line)
        (when @current-slide
          (swap! current-slide update :points conj
                 (str/trim (str/replace line #"^\s*[-*]\s+" ""))))

        ;; Additional content that's not empty
        (and (not (str/blank? line)) @current-slide)
        (swap! current-slide update :points conj (str/trim line))))

    ;; Add the last slide
    (when @current-slide
      (swap! slides conj @current-slide))

    @slides))

(defn generate-slide-content
  "Generate enhanced content for a slide using an LLM"
  [ollama-url slide]
  (let [prompt (str "Create professional PowerPoint slide content for a slide with the title: \""
                    (:title slide)
                    "\". Use these bullet points as guidance: "
                    (str/join ", " (:points slide))
                    ". Format your response as a JSON object with fields: 'title' (string), 'content' (array of bullet points), 'notes' (string, presenter notes), and 'visual_description' (string, what image or diagram would enhance this slide).")
        response (pyjama/ollama
                       ollama-url
                       :generate
                       {
                        :format {:type       "object"
                                 :required   [:title :content :notes :visual_description]
                                 :properties {:title   {:type "string"} :visual_description {:type "string"} :notes {:type "string"}
                                              :content {
                                                        :type  "array"
                                                        :items {:required   ["point"]
                                                                :properties {:point {:type "string"}}}}}}
                        :model  "llama3.1"
                        :prompt prompt} :response)]

    ;; Parse the JSON from the response using Cheshire
    ;; Handle potential issues with JSON formatting by extracting JSON from markdown code blocks if needed
    (try
      (let [json-str (if (re-find #"```json" response)
                       (second (re-find #"```json\s*([\s\S]*?)\s*```" response))
                       response)]
        (json/parse-string json-str true))                  ;; 'true' for keywordizing keys
      (catch Exception e
        ;; Fallback if JSON parsing fails
        (println "Warning: Failed to parse JSON response. Using raw slide data.")
        {:title              (:title slide)
         :content            (:points slide)
         :notes              "No notes generated."
         :visual_description "No visual suggestion."}))))

(defn create-title-slide
  "Create a title slide for the presentation"
  [ppt presentation-title subtitle]
  (let [master (.get (.getSlideMasters ppt) 0)
        layout (first (.getSlideLayouts master))            ; Title slide layout
        slide (.createSlide ppt layout)
        title-shape (.getPlaceholder slide 0)
        subtitle-shape (.getPlaceholder slide 1)]

    ;; Set title
    (.setText title-shape presentation-title)

    ;; Set subtitle
    (.setText subtitle-shape subtitle)

    slide))

(defn create-content-slide
  "Create a content slide with title and bullet points"
  [ppt slide-data]
  (let [master (.get (.getSlideMasters ppt) 0)
        layout (second (.getSlideLayouts master))           ; Title and content layout
        slide (.createSlide ppt layout)
        title-shape (.getPlaceholder slide 0)
        content-shape (.getPlaceholder slide 1)]

    ;; Set title
    (.setText title-shape (:title slide-data))

    (doseq [point (:content slide-data)]
      (println (:title slide-data) ">" point)
      (let [paragraph (.addNewTextParagraph content-shape)
            text-run (.addNewTextRun paragraph)]
        (.setIndentLevel paragraph 1)                       ;; Use setLevel instead of setIndentLevel
        (.setBullet paragraph true)
        (.setText text-run (:point point))))

    ;;; Add notes if available
    ;(when-let [notes (:notes slide-data)]
    ;  (let [notes-page (.getNotesPage slide)
    ;        notes-shape (.getPlaceholder notes-page 1)]
    ;    (.setText notes-shape notes)))
    ;
    ;;; Add visual description as a note (in a real app, you'd use this to generate/place images)
    ;(when-let [visual (:visual_description slide-data)]
    ;  (let [notes-page (.getNotesPage slide)
    ;        shape (.createTextBox notes-page)]
    ;    (.setText shape (str "Suggested visual: " visual))
    ;    (.setAnchor shape (Rectangle. 50 300 600 100))))

    slide))

(defn generate-pptx
  "Generate a complete PowerPoint presentation using the parsed outline"
  [ollama-url outline output-path]
  (let [slides (parse-outline outline)
        enhanced-slides (doall (mapv #(generate-slide-content ollama-url %) slides))
        ppt (XMLSlideShow.)]

    ;; Create title slide (using first slide's title as presentation title)
    (create-title-slide ppt
                        (or (:title (first enhanced-slides)) "Presentation")
                        "Generated with Clojure and Ollama")

    ;; Create content slides
    (doseq [slide enhanced-slides]
      (create-content-slide ppt slide))

    ;; Save the presentation
    (with-open [out (FileOutputStream. output-path)]
      (.write ppt out))

    (println "Successfully generated PowerPoint presentation with" (count enhanced-slides) "slides")
    (println "Output saved to:" output-path)

    ;; For debugging, we can save the enhanced slide data as JSON
    (spit (str output-path ".json") (json/generate-string enhanced-slides {:pretty true}))

    ;; Return the enhanced slides data for potential further processing
    enhanced-slides))

(defn -main
  "Main entry point for the application"
  [& args]
  (let [input-file (or (first args) "outline.txt")
        output-file (or (second args) "presentation.pptx")
        ollama-url (or (System/getenv "OLLAMA_URL") "http://localhost:11434/")]

    (println "Reading outline from:" input-file)
    (println "Using Ollama API at:" ollama-url)

    (if (.exists (io/file input-file))
      (let [outline (slurp input-file)
            result (generate-pptx ollama-url outline output-file)]
        (println "Successfully generated presentation with" (count result) "slides")
        (println "Output saved to:" output-file))
      (println "Error: Input file not found:" input-file))))

;; Example usage:
;; (generate-pptx "http://localhost:11434/api"
;;   "# Introduction\n- Project overview\n- Key objectives\n\n# Market Analysis\n- Current trends\n- Competitor landscape"
;;   "presentation.pptx")