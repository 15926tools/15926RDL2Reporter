# RDLReporter

This is a generic report writer. In the report tab, enter a Sparql query and endpoint. This query may have one to many relationships in it. In the report this is taken care of by making multiple columns for such relationships. 

## Usage

**Important:**

1. The report Sparql variables must be the same as selected in the query.
2. The primary key is the value that changes per single output record. 
3. It must be the first column and there must be a “yes” in the primary key setting .

**Note:**  If you do not like it that it must be the first column, the output is Excel, so just move the columns after reporting.

## Program Parameters
1. Sparql Endpoint: _Row 2B_
2. Sparql Query: _Row 3B_

**Query Example:**

```Sparql
prefix dm: <http://data.15926.org/dm/>
prefix chifos: <http://data.15926.org/cfihos/>
prefix skos: <http://www.w3.org/2004/02/skos/core#>

select 
    ?fklabel ?fkid ?docmasterlabel ?docmasterid ?def ?fktypeid ?subprop
    {
        ?docmasterid rdfs:label ?docmasterlabel .
        ?fkid rdf:type ?docmasterid .
        ?fkid rdfs:label ?fklabel .
        ?fkid skos:definition ?def .
        ?fkid rdf:type ?fktypeid .
        ?subprop rdfs:subPropertyOf ?fkid .
        ?fkid rdf:type dm:ClassOfClassOfInformationRepresentation .
    }
order by 
    ?docmasterlabel ?fklabel ?def ?fktypeid
```

**Configure Result Set**

![alt text](https://https://github.com/15926tools/15926RDL2Reporter/resultset_config.JPG "Result set Config")

1. Sparql Variable - The binding to which the variable will assign data to.
2. Place multiples in max how many columns - Addressing the carthesian product that will result from running a select query.
3. Is primary key - Used to identify which sparql variable will act as the record unique idenitifer.
4. Column header - Used for mapping the result bindings to a column on the result tab sheet.
5. Column Width - Result tab column width.


**Not Supported**

Construct Queries
