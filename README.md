# Web-Automation-using-Excel-VBA

A NOC automation tool used to find customer ID's and prevent the repetitive task of copying and pasting data from Excel to a web form. 

Data is extracted from Excel and populated into Kibana Logstash website text fields (must be in Internet Explorer). Kibana returns the resulting HTML page. The HTML is scraped and parsed to find customer ID's. The customer ID's are inserted back into Excel in the L column
