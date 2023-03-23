# indi-go

A SANS/GIAC Index builder written in `golang`. This isn't anything super robust, we take an excel spreadsheet with sheets for each book and iterate through the sheets to make markdown files for each book in your index. My index consists of `Page Number`, `Slide Title`, and two or three keyword columns that I can look back on. I might update this to create maps for each page and make a aggregated index but it is not happening before my next test.

## converting to PDF

I use npm's md-to-pdf but you can use whatever you like. I am writing this down because I keep forgetting "what I like" https://www.npmjs.com/package/md-to-pdf


