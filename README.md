# pubmed2doc

## Introduction

Would you like your PubMed search results to be saved as PDF or Ms Word file format? If yes, then look 
no further, I present to you **pubmed2doc**, a python module implemented to write your PubMed search
results to PDF or Word file. The basic workflow of the module in a nutshell is as follow:
1. It takes your query as an input and searches it against the PubMed database. 
2. The returned PubMed results will be then prepared in one of the two display options i.e citation or listview (based 
on the user's choice).
3. The prepared results will finally be written to user's supplied file choice (PDF or Word).

To my knowledge, PubMed doesn't offer this feature i.e saving (or exporting) the search results to PDF or Word and 
therefore this module provides the following advantages:

1. It doesn't require any intermediate software to aid in exporting the PubMed results to PDF or Word. 
2. It Keeps the document stored locally and can be retrieved at a later time. 
2. The generated document is easier to read, view and edit comparing with the csv or text files exported by PubMed. 
3. It can also be ported or incorporated easily into the user's own research resources.

## Getting Started

* You need to have python version >= 3.4 installed and a terminal (or similar program) to execute 
the module through the command line. 


* The following sections will guide you to execute the module successfully in a step-by-step fashioned.


## Prerequisites

* The following are the core Python libraries (dependencies) required for this module:

  ```bash
  biopython
  fpdf
  python-docx
  ```

## Install 
1.  First, you need to download the compressed file in GitHub.

2.  Uncompress it and change to the directory ```cd```:
      ```bash
    unzip Pubmed_To_Docs.zip
    cd Pubmed_To_Docs/
      ```    
3. It is highly recommended to create conda or python virtual environment and 
install the aforementioned dependencies within such environment (OPTIONAL):
   * Creating conda virtual environment, please change "your_env_name" your desire environment name:
 
        ```bash
        conda create -n your_env_name
        conda activate your_env_name
       ```
   
   * Creating python virtual environment:
 
       ```bash
       pip install virtualenv
       virtualenv your_env_name
       source your_env_name/bin/activate
       ```
   
   * **Note:** don't forget to de-activate your environment once you are completely done running the module:
       ```bash
       # for conda env
       conda deactivate
       
       # for python env
       deactivate
       ```

4. You may now try to install the dependencies by running the following command:
    ```bash
    sudo pip install -r requirements.txt
    ```
- You will be prompted to enter a password if you run the  ```sudo``` command, the password is
the root password (system administrator) of the computer you are currently running.

## Running the Module (pubmed2doc.py)

* Please make sure you are in the same directory location as the module pubmed2doc.py, and you may try to 
check the presence of the module by running the following command:
    ```bash
    ls -ltrh *.py  

* #### Inputs + Options

    + -q: a query to be searched against PubMed database (REQUIRED)

    + -e: a user's email to access to the PubMed db (REQUIRED)

    + -pdf: write PubMed results to PDF (OPTIONAL) (Default value is True)

    + -word: write PubMed results to Word (OPTIONAL) (Default value is False)

    + -retmax: total num of records from query to be retrieved (OPTIONAL)
    (Default value is 20)

    + -sopt: type of returned results display options (citation or listview)
    (OPTIONAL) (Default is citation)
  

* #### Output
    + a Word or PDF document consists of the results from PubMed according to the type of display option 
    chosen by the user
  

* ### Command Example
    ```bash
    python pubmed2doc.py \
    -q "gene expression" \
    -e "your_email" \
    -pdf F \
    -word T \
    -sopt listview
    ```
* The above command is just an example, you can create your own advance query and include other arguments, please see
above sub-section "Inputs + Options"

* Please note that the output file will be saved in a directory called "output"

* That is it! Please check the output folder for your file

## Asking for Help

If you have an issue (error or bug) when executing the module or have any other enquires please feel free to send 
me an [email](mailto:nawafalomran@hotmail.com) and I will reply as soon as I can