(ns waves.cli
  (:gen-class)
  (:require [clojure.tools.cli]
            [waves.core]))

(def cli-options
  [["-h" "--help"]
   ["-i" "--input INPUT" "Input PowerPoint file (default: input.pptx)"
    :default "input.pptx"]
   ["-o" "--output OUTPUT" "Output PowerPoint file (default: output.pptx)"
    :default "output.pptx"]
   ["-m" "--model MODEL" "Model to use (default: llama3.2)"
    :default "llama3.2"]
   ["-u" "--url URL" "Server URL (default: http://localhost:11434)"
    :default "http://localhost:11434"]
   ["-s" "--system PROMPT" "System prompt (default: translation prompt)"
    :default "You are a machine translator from Japanese to English. You are given one string each time. When the string is in English you return the string as is. If the string is in Japanese, you answer with the best translation. Your answer will only contain the translation and the translation only, nothing else, no question, no explanation. If you do not know, answer with the same string as input"]
   ["-p" "--prompt-template PROMPT-TEMPLATE" "Prompt template (default: %s)"
    :default "Translate the following: %s"]
   ["-d" "--debug" "Enable debug mode (default: false)"
    :default true
    :parse-fn #(Boolean/valueOf ^String %)]])

(defn -main [& args]
  (let [{:keys [options errors summary]} (clojure.tools.cli/parse-opts args cli-options)]
    (cond
      (:help options)
      {:exit-message (println "\nUsage:\n" summary) :ok? true}
      (seq errors)
      (do
        (println "Error parsing options:")
        (doseq [err errors] (println "  " err))
        (println "\nUsage:\n" summary)
        (System/exit 1))
      :else
      (let [
            {:keys [^String input ^String output] :as config} options
            state (atom options)
            ]
        (swap! state assoc :stream false :file-path input)
        (clojure.pprint/pprint options)
        (waves.core/update-ppt-text state)))))