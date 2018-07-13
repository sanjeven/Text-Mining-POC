# Text mining POC

This POC uses python and excel to collect data from wiki like sites for further analysis.

<b>Requirements</b>
<ul>
<li type = "square">Python2.7</li>
<li type = "square">Beautifulsoup4</li>
<li type = "square">Pandas</li>
<li type = "square">win32com.client</li>
<li type = "square">itertools</li>
<li type = "square">Excel</li>
<li type = "square">Copy of the excel macro provided</li>
 </ul>
<br>
<b>How to Run</b><br>
<i>python *scriptname* -s *start year* -e *end year*</i>
<br></br>
<ul>
<li type = "square">The excel macro combines 2 columns together as wiki's html is not standardised which meant mining had to be done for 2 criterias.</li>
<li type = "square">To run the script without the excel macro, simply comment out server.run_author_macro()</li>
<li type = "square">*question marks* that appear in the excel document are caused by special characters that are not able to be captured using the encoding method (not yet fixed)</li>
</ul>
