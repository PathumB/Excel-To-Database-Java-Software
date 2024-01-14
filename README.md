#### What is this?
I previously worked for a web development company where we created a large-scale web application for a client. The client had been maintaining customer details in an Excel sheet, which they provided to us after the web application was developed. Our task was to transfer this data to the database, hosted on AWS.

Our backend engineer instructed me to design a software solution capable of uploading the text-based data from the Excel sheet to a local MySQL database. Additionally, some rows in the Excel sheet contained Google Drive links, and the corresponding documents needed to be downloaded and uploaded to an AWS S3 bucket.

Here is a high-level overview of the process:

- The software is designed to run on both Windows and various Linux distributions.
- It requires Java Runtime Environment (JRE) for execution.
- The software operates through the command line interface (CLI).
  
The successful implementation of this solution allows for efficient data transfer from the Excel sheet to the local MySQL database and handles the download and upload of documents from Google Drive links to the AWS S3 bucket
