(ns ppt.core
  (:import (org.apache.poi.xslf.usermodel XMLSlideShow XSLFTable XSLFTableRow XSLFTableCell XSLFTextParagraph)
           (org.apache.poi.sl.usermodel VerticalAlignment TableCell$BorderEdge TextParagraph$TextAlign Insets2D)
           (java.awt Rectangle Color)))


(defn add-table-basic!
  "Render a 2D vector/seq `table` onto `slide` at rect [x y w h].
   Options:
     :header?     true/false (default true)
     :font-family string     (default \"Inter\")
     :header-fill java.awt.Color (default light gray)
     :body-size   double     (default 11.0)
     :header-size double     (default 12.0)
     :pad         Insets2D   (default (Insets2D. 2 6 2 6))"
  [slide table x y w h & [{:keys [header? font-family header-fill body-size header-size pad]
                           :or {header? true
                                font-family "Inter"
                                header-fill (Color. 245 245 245)
                                body-size 11.0
                                header-size 12.0
                                pad (Insets2D. 2.0 6.0 2.0 6.0)}}]]
  (when (and (seq table) (seq (first table)))
    (let [data (mapv vec table)                         ;; ensure associative indexing
          rows (count data)
          cols (count (first data))
          ^XSLFTable t (.createTable slide)]
      (.setAnchor t (Rectangle. (int x) (int y) (int w) (int h)))

      ;; Build rows & cells; set text exactly once per cell.
      (dotimes [r rows]
        (let [^XSLFTableRow row (.addRow t)]
          (.setHeight row (/ (double h) rows))
          (dotimes [c cols]
            (let [^XSLFTableCell cell (.addCell row)
                  v (str (get-in data [r c] ""))]       ;; safe indexing
              (.setText cell v)))))                     ;; creates one paragraph + one run

      ;; Set column widths after columns exist.
      (dotimes [c cols]
        (.setColumnWidth t c (/ (double w) cols)))

      ;; Style pass on the run that already contains text.
      (dotimes [r rows]
        (dotimes [c cols]
          (let [^XSLFTableCell cell (.getCell t r c)
                ^XSLFTextParagraph p (first (.getTextParagraphs cell))
                run (first (.getTextRuns p))]
            (.setVerticalAlignment cell VerticalAlignment/MIDDLE)
            (.setInsets cell pad)
            (.setWordWrap cell true)
            (.setFontFamily run font-family)
            (.setFontSize run (if (and header? (zero? r)) header-size body-size))
            (when (and header? (zero? r))
              (.setTextAlign p TextParagraph$TextAlign/CENTER)
              (.setBold run true)
              (.setFillColor cell header-fill))
            ;; light grid borders (lowercased edges per your POI)
            (doseq [edge [TableCell$BorderEdge/left TableCell$BorderEdge/right
                          TableCell$BorderEdge/top  TableCell$BorderEdge/bottom]]
              (.setBorderColor cell edge (Color. 200 200 200))
              (.setBorderWidth cell edge 1.0)))))

      t)))

(defn add-tables-paginated!
  "Lay tables top→bottom; when the next table won't fit, start a new slide."
  [^XMLSlideShow ppt tables
   {:keys [x y0 w h gutter top-margin bottom-margin] :as opts}]
  (let [page (.. ppt getPageSize)
        page-w (.getWidth page)
        page-h (.getHeight page)
        x      (or x 60)
        y0     (or y0 120)
        w      (or w (- page-w 120))        ; leave some side padding by default
        h      (or h 160)
        gutter (or gutter 20)
        top-margin    (or top-margin y0)
        bottom-margin (or bottom-margin 40)
        y-max (- page-h bottom-margin)]
    (loop [remaining tables
           slide (.createSlide ppt)
           y y0]
      (when (seq remaining)
        (let [tbl (first remaining)
              next-y (+ y h)]
          (if (> next-y y-max)
            ;; new slide
            (recur remaining (.createSlide ppt) top-margin)
            ;; place table here
            (do
              (add-table-basic! slide tbl x y w h opts)
              (recur (rest remaining) slide (+ y h gutter)))))))))

(defn add-tables-one-per-slide!
  "Each table gets its own slide at a fixed rect."
  [^XMLSlideShow ppt tables & [opts]]
  (doseq [tbl tables]
    (let [slide (.createSlide ppt)]
      (add-table-basic! slide tbl 60 120 600 180 opts))))

(defn ^:private empty-table? [t]
  (or (nil? t) (not (seq t)) (not (seq (first t)))))

;
;

;; matches  --- ,  :--- ,  ---: ,  :---:  (with optional spaces/pipes)
(def ^:private md-sep-cell-re #"^\s*\|?\s*:?-{3,}:?\s*\|?\s*$")

(defn ^:private md-separator-row?
  "True if every cell in the row is a Markdown header separator."
  [row]
  (and (seq row)
       (every? #(re-matches md-sep-cell-re (str %)) row)))

(defn strip-second-md-separator-row
  "If the 2nd row is a Markdown header-separator (--- / :---:), drop it."
  [table]
  (let [data (mapv vec table)]
    (if (and (> (count data) 1)
             (md-separator-row? (data 1)))
      ;; keep row 0, drop row 1, keep the rest
      (vec (concat [(data 0)] (subvec data 2)))
      data)))
;
;

(defn add-first-table-to-slide!
  "Place the first non-empty table from `tables` onto `slide` at [x y w h].
   Returns {:placed? boolean, :remaining seq-of-tables}."
  [slide tables {:keys [x y w h] :or {x 60 y 120 w 600 h 180}} & [style-opts]]
  (let [[_skipped rest-tables] (split-with empty-table? tables)]
    (if-let [tbl (first rest-tables)]
      (let [tbl* (strip-second-md-separator-row tbl)]
        (add-table-basic! slide tbl* x y w h style-opts)
        {:placed? true
         :remaining (rest rest-tables)})
      {:placed? false
       :remaining tables})))