#!/usr/bin/env bash
: '
this shell script was created to run pubmed2doc with tweaking some of
its arguments.
'

# You should insert your email below.
email="INSERT_YOUR_EMAIL_HERE"
# you can change the query to your own query
query="gene therapy in cancer"

# write the results to PDF with a citation view and no abstract
python pubmed2doc.py \
    -q "$query" \
    -e $email \
    -pdf T \
    -word F \
    -sopt citation \
    -abs F

# write the results to PDF with a listview and no abstract
python pubmed2doc.py \
    -q "$query" \
    -e $email \
    -pdf T \
    -word F \
    -sopt listview \
    -abs F

# write the results to Word with a listview and no abstract
python pubmed2doc.py \
    -q "$query" \
    -e $email \
    -pdf F \
    -word T \
    -sopt listview \
    -abs F

# write the results to Word with a citation and no abstract
python pubmed2doc.py \
    -q "$query" \
    -e $email \
    -pdf F \
    -word T \
    -sopt citation \
    -abs F