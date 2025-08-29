(ns md.core
  (:require [clojure.java.io :as io]
            [clojure.string :as str]))

(defn parse-md [filepath]
  (with-open [rdr (io/reader filepath)]
    (let [lines (line-seq rdr)]
      (loop [remaining lines
             slides []
             current nil
             in-code? false
             code-block []
             ]
        (if (empty? remaining)
          (if current (conj slides current) slides)
          (let [line (first remaining)
                header-match (re-find #"^(#+) (.+)$" line)
                img-match (re-find (re-pattern "!\\[.*?\\]\\((.*?)\\)") line)
                is-table (re-matches #"\|.*\|" line)
                is-code-start (re-matches #"^```(\w+)?$" line)
                is-code-end (re-matches #"^```$" line)]

            (cond
              ;; start code block
              (and is-code-start (not in-code?))
              (recur (rest remaining) slides current true [])

              ;; end code block
              (and is-code-end in-code?)
              (recur (rest remaining)
                     slides
                     (update current :code-blocks (fnil conj []) {:lang "code" :code code-block})
                     false
                     [])

              ;; inside code block
              in-code?
              (recur (rest remaining) slides current true (conj code-block line))

              ;; new header
              header-match
              (let [[_ hashes title] header-match
                    level (count hashes)]
                (recur (rest remaining)
                       (if current (conj slides current) slides)
                       {:type        (if (= level 1) :title :slide)
                        :title       title
                        :level       level
                        :body        []
                        :tables      []
                        :images      []
                        :code-blocks []}
                       false
                       []))

              ;; image
              img-match
              (let [img-path (second img-match)]
                (recur (rest remaining)
                       slides
                       (update current :images conj img-path)
                       false
                       []))

              ;; table
              is-table
              (let [[table-lines rest-lines] (split-with #(re-matches #"\|.*\|" %) remaining)
                    rows (->> table-lines
                              (map #(->> (str/split % #"\|")
                                         (map str/trim)
                                         (remove str/blank?))))]
                (recur rest-lines
                       slides
                       (update current :tables conj rows)
                       false
                       []))

              ;; normal line
              :else
              (recur (rest remaining)
                     slides
                     (update current :body conj line)
                     false
                     []))))))))
