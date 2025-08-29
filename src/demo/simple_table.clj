(ns demo.simple-table
  (:import (org.apache.poi.xslf.usermodel XMLSlideShow)
           (java.io FileOutputStream)
           (java.awt Rectangle)))

(defn random-cell-text []
  (subs (str (java.util.UUID/randomUUID)) 0 8)) ; short random token

(defn add-table-very-simple!
  "Builds a rows×cols table with plain text. No styling, no extra runs."
  [slide rows cols x y w h]
  (let [t (.createTable slide)]
    (.setAnchor t (Rectangle. (int x) (int y) (int w) (int h)))

    ;; create rows & cells, set text exactly once per cell
    (dotimes [r rows]
      (let [row (.addRow t)]
        (.setHeight row (/ (double h) rows))
        (dotimes [c cols]
          (let [cell (.addCell row)]
            (.setText cell (str "R" r "C" c " " (random-cell-text)))))))

    ;; set column widths *after* first row created (columns now exist)
    (dotimes [c cols]
      (.setColumnWidth t c (/ (double w) cols)))

    t))

(defn -main [& _]
  (let [ppt (XMLSlideShow.)
        slide (.createSlide ppt)]
    (add-table-very-simple! slide 3 4 60 120 600 180)
    (with-open [out (FileOutputStream. "simple-table.pptx")]
      (.write ppt out))
    (println "✅ Wrote simple-table.pptx")))

