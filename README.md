## What is this?

This a simple demonstration of how you can automate a report using the data of a retail company. The idea is that you can access the database on a regular basis (for example once every month), get the data you need, create the necessary tables, write them in an excel file and send the file to the client.

This python project is the first step towards that as we use a fixed xlsx file to demonstrate the process. 

## What are the main steps?

- Create a virtual environment and install the requirements.txt
- Run the main.py script
- Open the output.xlsx that is generated on the same directory

## Research questions

*  Which type of customer group has spent the most money?
    * How much money has been spent on each type of customer group on refunded orders?

    * Which type of product drove the most sales for customer_groups [A - D] where product value was lower than 100?

* Which product has been sold / refunded the most?
   * How many products with value more than 450 have been refunded for customer_group [F - J]? 
   * What is the value of these orders?

* Is it true that the higher the type of product, the more money has been spent on it?
   * Can you plot the ratio between the sum of prices and the number of orders made for each product type?
    
* Which product group is the most popular for customer groups [A, B, C, D]?
   * Which customer group buys the most of each product group?
    
* What is the probability that an order by a customer group [A - D] is refunded?



## Future imporvements

* This is a quick demonstration of the functionality of the automated report. More in depth work needs to be done in order to evalute  the research questions.
* Spin up a dashboard on Streamlit to create visualisations. An excel as a deliverable might be sufficient for a user tha can analyse the data.
* Build an infrastructure on AWS so that you can upload the data there on an S3 bucket and then trigger the python script to create the deliverable.



## Owner:

Vasilis Panagaris