(ns demo.styled-table
  (:require [clojure.pprint]
            [ppt.core])
  (:import (java.util UUID)
           (org.apache.poi.xslf.usermodel XMLSlideShow)
           (java.io FileOutputStream)))

(defn -main [& _]
  (let [ppt (XMLSlideShow.)
        slide (.createSlide ppt)
        table (for [r (range 3)]
                (for [c (range 4)]
                  (str (if (zero? r) (str "Col " (inc c)) (str "R" r "C" c))
                       " – "
                       (subs (str (UUID/randomUUID)) 0 6))))]
    (ppt.core/add-table-basic! slide table 60 120 600 180)
    (with-open [out (FileOutputStream. "styled-table.pptx")]
      (.write ppt out))
    (println "✅ Wrote styled-table.pptx")))
