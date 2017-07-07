# summer-stretch
Python scripts that automatically generate progress reports for a roster of students using a base template and a file of options.

### Getting Started
* Install Python v3: https://www.python.org/downloads/
* Add C:\Python and C:\Python\Scripts as a paths in environment variables
* Install the python-docx package
```
pip install python-docx
```
* Clone the repo https://github.com/ezhai24/summer-stretch.git
* For progress reports: create custom `roster.txt`, `improvements.txt`, and `template.docx` files and place them in a common directory
* For final transcripts: create custom `roster.txt`, `comments.txt`, and `template.docx` files and place them in a common directory
* In command prompt, navigate to the common directory and run
```
# for progress reports
python progress_reports.py /path/to/input/files progress-report-#

# for final transcripts
python final_transcripts.py /path/to/input/files
```

### The Roster File
Each line in this file corresponds to a student (and their respective report). The parameters should be separated by a space and in the order as follows for progress reports:
* Gender Pronouns: M(ale), F(emale), N(eutral)
* First name
* Last name
* Improvement code: the first two letters of the corresponding in `improvements.txt`, capitalilzed
* Midterm grade (optional)

For final transcripts, the order is as follows:
* Gender Pronouns: M(ale), F(emale), N(eutral)
* First name
* Last name
* Final percent grade
* Whether the student should recieve high school credit for the course: Y(es), N(o)

### The Improvements File
Each line in this file corresponds to an area of improvement (preparation, content knowledge, etc.) that students believed they've made the most progress in.

### The Comments File
There are two lines in this file. The first is the comment for students scoring higher than 80 and the second is for those scoring 80 or lower. Each line should be followed by a blank line.

### The Template Document
This is a word document that serves as a template for the reports. Progress reports should be named `Template_{Progress Report #}` and the final transcript template should be named `Template_FinalTranscript`.

### Replacements
These are strings that can be used in the template document. The program will replace them with the values defined in `roster.txt` and `improvements.txt`.
* &lt;firstname>
* &lt;lastname>
* &lt;spn>: subject pronoun (he, she, they)
* &lt;opn>: object pronoun (him, her, them)
* &lt;ppn>: possessive pronoun (his, hers, their)

Only for progress reports:
* &lt;improvement>
* &lt;testscore>

Only for final transcripts:
* &lt;percentgrade>
* &lt;lettergrade>
* &lt;hscredit>
* &lt;comment>