# NFS-e Automatic
![telaprincipal](https://user-images.githubusercontent.com/54956673/143236506-8c7624b7-0fa5-4d9e-a31a-58af31ade0e3.PNG)

Script for automation of NFS-e releases (Brazilian electronic invoice standard) in the Fiscal Writing
module of the Thomson Reuters Domínio ERP system.

> Status: Developing ⚠️


<h2> About Project </h2>
<p> The project aims to eliminate as much of the manual work as possible to enter the data of the electronic service invoices provided to the companies
and sent to the client in PDF format. This work requires the fiscal analyst to spend very large on reading the data and inserting them one by one within 
the tax bookkeeping system for the appropriate measures with the agencies that oversee the accounting and tax activities of Brazilian companies.
With this script, it was possible to reduce between 70 and 80% the need to enter data that would be performed manually by the fiscal analyst and noticed a 
significant reduction of time to complete the activity. Work that was done for days was reduced to just a few minutes after the dataframe was created
to iterate the data into the system to which the script was proposed. </p>

## Technologies Used

<table>
<tr>
 <td>Python</td>
 <td>Pandas</td>
 <td>Pyautogui</td> 
</tr>
<tr>
 <td>3.10.0</td>
 <td>1.3.4</td>
 <td>0.9.53</td>
</tr>
</table>

<p> The most laborious part that would be the extraction of pdf data to create the Dataframe, we use the Optical Character Recognition (OCR) technology provided 
by the Rossum cloud application (https://rossum.ai/) that allows the reading and extraction of data from invoices, bank statements and various other document models using artificial intelligence, standardizing the extraction from the identification of the fields within the PDF files that will be used in the creation of our DataFrame and insertion of the data in the tax bookkeeping system with the script developed in Python 3.10.0.</p>

## How to run the application

* create virtual enviroment;
* execute command 'pip install -r requirements.txt' for install the projects dependencys.
