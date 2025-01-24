# Dashboard-for-an-Online-Boutique
1 Introduction
This application was developed to support the operations of a boutique e-commerce business, focusing on efficient order management and streamlined processes. The system integrates an Access database with an Excel-based front end, offering functionalities such as data visualization, order tracking, and seamless data entry. The purpose of this project is to provide a practical tool for handling orders, while demonstrating how VBA can bridge the gap between databases and user-friendly interfaces. The data used for this project originates from a publicly available dataset on Brazilian e-commerce, which includes extensive order and customer information.

2 Database
The database used in this project is designed to store detailed information about orders. It consists of three tables ‘Orders’, ‘Products’, and ‘Customers’, and is carefully structured to capture key aspects of an order’s lifecycle, from placement to delivery.
 
 
The ‘Orders’ table contains: OrderId, OrderStatus, PurchaseTime, EstimatedDeliveryTime, DeliveredTime, ProductId, NumberOfProducts, CustomerId. It records essential information of each order, enabling further statistic analysis.
 
The ‘Products’ table contains: ProductId, ProductPrice, ProductWeight.
 
The ‘Customer’ table contains: CustomerId, PhoneNumber, Age.
To support data retrieval for the application, three queries were developed to extract key order details: 
 
‘CalculateDelayDuration’: This query determines whether an order was delayed and calculates the delay duration for those that were.
 
‘Top3HeaviestOrders’: This query identifies the top three heaviest orders. This feature can be extended to explore the relationship between package weight and delivery delays, offering valuable insights into reducing delays.
 
‘Top3HighestTotalPrice’: This query retrieves the orders with the three highest total prices. Such analysis provides a straightforward overview of order price scales and insights into customer groups with higher spending habits.

3 Front-End
The Excel workbook serves as the front end for this application and is structured with three key worksheets. 
 
The first sheet ‘Orders’ displays all order details. This sheet includes color-coded status indicators to help users quickly identify the progress of each order. 
The second sheet is a Sales Summary, featuring a pivot table that aggregates sales data by product. This provides valuable insights into product performance. 
The third sheet ‘NewOrder’ is designed for data entry, allowing users to input new orders directly. Built-in validation ensures that all data entered adheres to the required format and standards, minimizing the risk of errors.

4 VBA Middleware
The functionality of this application is powered by VBA code, which acts as middleware to connect the Access database with the Excel interface. Several subroutines were developed to achieve specific tasks.
The ‘GetOrdersData’ subroutine retrieves order details from the database and populates the Orders sheet. This subroutine also invokes the ‘ColorCodeOrderStatus’ procedure, which applies conditional formatting to visually differentiate order statuses. 
Another subroutine, ‘InsertNewOrder’, enables users to add new orders directly from the Order Form sheet. This procedure includes error handling to prevent issues such as data type mismatches and ensures that only valid data is inserted into the database.

5 Conclusion
This application demonstrates how a small-scale tool can address real business needs in an efficient manner. By integrating an Access database with Excel and leveraging VBA, the system provides a practical solution for order management. To scale this application for a real-world business, additional features could be implemented, such as support for multiple tables to capture customer and product details, dynamic dashboards for real-time analytics, and a web-based front end for multi-user access.
The source code and associated files for this project are available on GitHub, providing further exploration or adaptation.

6 References
The data used for this project was sourced from Kaggle’s Brazilian E-Commerce dataset: https://www.kaggle.com/datasets/olistbr/brazilian-ecommerce. 
Additional resources include lecture materials and examples provided during coursework.

