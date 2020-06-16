# RDLReporter
===============================================
This is a generic report writer. In the report tab, enter a Sparql query and endpoint. This query may have one to many relationships in it. In the report this is taken care of by making multiple columns for such relationships. For example, if an object has 3 rdf:type relationships to classes, you can put that in 3 columns titled “type”.
Important:
•         The report Sparql variables must be the same as selected in the query
•         The primary key is the value that changes per single output record. It must be the first column and there must be a “yes” in the primary key setting
•          If you do not like it that it must be the first column, the output is Excel, so just move the columns after reporting
