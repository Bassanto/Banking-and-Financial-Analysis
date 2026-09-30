#  Nigerian Banking Operations Dasboard

Here is the file for the full interactive dashboard [View Here](./Naijafin_Analysis%20(Recovered).xlsx)

 
## Project Overview
 
This project analyzes transaction activity, customer behavior, revenue generation, branch performance, and operational efficiency within a Nigerian banking environment. Using a combination of Excel functions, Pivot Tables, and Pivot Charts, I analyzed the dataset to generate actionable insights, quantify performance metrics, and address key business objectives.
 
## Business Questions
- What were total transaction value, fee income and successful transaction count in FY2025?
- Which branches generated the highest transaction value and fee income?
- Which product categories and products contributed most to revenue/fees?
- How did transaction value and fee income trend by month?
- What is the success/failure/reversal rate by channel?
- How does customer risk segment relate to transaction activity and value?
- What are total monthly payroll costs by department?
- Identify data-quality issues and quantify how many records required correction.
- Produce an executive dashboard using PivotTables, PivotCharts, slicers and KPI cards.

## Dataset
1. `customer_raw`
This contains information customers

                              
- `Customer_ID`        | Unique identifier for each customer and the primary customer key. Used to link customers to their accounts and transactions.                       
- `Full_Name`          | Customer name. Primarily used for identification and record-level reference.                                                                                  
- `Gender`             | Customer gender. Used as a demographic attribute for potential segmentation.                                                                                                                                                   
- `Employment_Status`  | Customer employment category, such as Salaried, Self-Employed, Business Owner or Contract.                                                         
- `Monthly_Income_NGN` | Declared monthly income in Nigerian Naira. Used as a financial/customer attribute.                                                                 
- `Risk_Segment`       | Customer risk classification: Low, Medium or High. Used to analyse differences in transaction activity and transaction value across risk segments. 
- `Join_Date`          | Date the customer joined the bank. Used to calculate customer tenure and create tenure bands for behavioural analysis.                             


2. `account_raw`


- `Account_ID`          | Unique identifier for each bank account. Serves as the primary account key.                                                         
- `Customer_ID`         | Identifies the customer who owns the account. Used to establish the Customer → Account relationship.                                
- `Account_Type`        | Type/category of bank account held by the customer.                                                                                 
- `Opening_Date`        | Date the account was opened. Used to validate transaction dates and ensure transactions do not occur before the account was opened. 
                                                                      
- `Account_Status`      | Current status of the account.                                                                                                      
- `Branch_Code`         | Identifies the branch associated with the account and links the account to the branch lookup table.                                 
                  


3. `transaction_raw`

- `Transaction_ID`     | Unique transaction identifier. Used to identify and validate individual transactions; duplicate records were intentionally included in the raw data for data-quality checks. 
- `Account_ID`         | Identifies the account associated with the transaction. Links transactions to `account_raw`.                                                                                 
- `Customer_ID`        | Identifies the customer associated with the transaction. Used for customer-level analysis and relationship validation.                                                       
-`Transaction_Date`   | Original transaction date. The raw data contains deliberately mixed date formats and required cleaning before analysis.                                                      
- `Product_ID`         | Identifies the banking product associated with the transaction and links to the product lookup table.                                                                        
- `Product_Category`   | Broad classification of the transaction, such as Card, Cash, Bills, Lending or Transfer.                                                                                     
- `Product_Type`       | More specific transaction type within the product category, such as Purchase, Withdrawal, Deposit or Repayment.                                                              
- `Direction`          | Indicates the transaction direction, such as inflow or outflow.                                                                                                              
- `Amount_NGN`         | Monetary value of the transaction in Nigerian Naira. This is the principal measure used for transaction-value analysis.                                                      
- `Fee_NGN`            | Fee charged on the transaction. Used to analyse fee income.                                                                                                                  
- `Channel`            | Channel through which the transaction occurred, such as Mobile App, POS, ATM, USSD, Internet Banking or Branch.                                                              
- `Transaction_Status` | Indicates whether a transaction was Successful, Failed or Reversed. Used to distinguish valid transactions from unsuccessful activity.                                       
- `Branch_Code`        | Identifies the branch associated with the transaction and supports branch-level performance analysis.                                                                        

4. `branch_lookup`
- `Branch_Code` | Unique branch identifier and the key used to link branch information to other datasets.               
- `Branch_Name` | Name/location of the banking branch, such as Victoria Island, Ikeja, Garki or Ibadan.                 
- `State`       | Nigerian state in which the branch is located.                                                        
- `Region`      | Geographic region grouping, such as South West, South East, South South, North Central or North West. 


5. `employee_payroll`

- `Employee_ID`        | Unique identifier for each employee.                                                                        
- `Employee_Name`      | Employee name used for record-level identification.                                                         
- `Branch_Code`        | Identifies the branch where the employee is assigned and links employees to `branch_lookup`.                
- `Department`         | Employee's functional department, such as Finance, Retail Banking, Customer Service, Operations or Digital. 
- `Monthly_Salary_NGN` | Employee's monthly salary in Nigerian Naira. Provides the primary payroll cost measure.                     
- `Hire_Date`          | Date the employee joined the organisation.                                                                  
- `Employment_Status`  | Current employment status, such as Active or On Leave.                                                      

## Data Cleaning Process
The dataset came with inconsistencies in the letter format and also in the date column. They were also null values which I replaced using conditional avaerage or median value. 

- Date 

To prevent import issues, all date columns were initially loaded as text from Power Query. Since the source data contained inconsistent date formats, I applied the DATEVALUE() function to convert the text strings into valid Excel date values, ensuring accurate date recognition and calculations.

```
=IF([@[Transaction_Date]]="","",
IF(ISNUMBER(SEARCH("/",[@[Transaction_Date]])),
  IF(LEN([@[Transaction_Date]])=10,
    DATE(VALUE(RIGHT([@[Transaction_Date]],4)),
         VALUE(LEFT([@[Transaction_Date]],2)),
         VALUE(MID([@[Transaction_Date]],4,2))),
    DATE(2000+VALUE(RIGHT([@[Transaction_Date]],2)),
         VALUE(MID([@[Transaction_Date]],4,2)),
         VALUE(LEFT([@[Transaction_Date]],2)))
  ),
  IF(AND(ISNUMBER(VALUE(LEFT([@[Transaction_Date]],4))),
         MID([@[Transaction_Date]],5,1)="-"),
    DATE(VALUE(LEFT([@[Transaction_Date]],4)),
         VALUE(MID([@[Transaction_Date]],6,2)),
         VALUE(RIGHT([@[Transaction_Date]],2))),
    DATEVALUE([@[Transaction_Date]])
  )
))
```
I validated this by joining account opening date to the transaction date table  to ensure that the transaction date is after the account opening date.
 
 - Standardization of words

 Some column such `transaction_status`, `gender` etc has error in the formatting such as spaces, different case letters etc. I solved this by adding a new column and and using PROPER(TRIM( )) function to standardise the columns.This is to avoid repetition during analysis.

 - Null Values
 
The monthly_income column had several null values that were handled through conditional median imputation. The IF(MEDIAN()) logic was used to calculate and assign median values within specific categories, ensuring that replacements reflected the characteristics of each subgroup rather than the dataset as a whole. I choose median over average to avoid effect of outlier.

- Duplicates

Duplicate records were identified through column formatting and data inspection. After verifying that the entries were true duplicates, they were removed to improve data accuracy and consistency.

## Analysis and Insights
1. The transaction trend has a significant partner and that is increase during 
holiday period. January, ,May , July , August and December has a signifcant increase compared to the previous month.
2. Only 249 accounts are active out of 402 accounts which is not really a positive result. 
3. Ibadan branch has the least transaction value with over 88,000,000 naira lesser than Wuse branch which is also the highest branch. 
4. Mobile App and POS are the dorminant is the transaction channel with over 45% dominant over 4 other methods .
5. The analysis have shown that customers acquired in the last 4 months are transacting better than the older customers. Ths indicates that the bank's campaign program and operating system have improves in acquring active customers.
6. Customers with low risk segment has highest transaction and transaction value over medium and high risk segment.


## Recommendation
1. The bank should maintain their current banking operating system as the conversion rate of the last for month is higher than the previous years. 
2. The technicality of mobile app should be improved as it is the only channel that recorded failed and reversed transaction. 
3.  Launch a reactivation campaign targeting dormant and closed accounts 
, with only 249 of 402 accounts active, a substantial share of the 
customer base is currently generating no transaction value.
4. The `monthly_targets` file should be investigated further . H2 2025 
targets (₦11.08B) are roughly 12x higher than both the H1-based 
naive projection (₦856M) and the actual H2 transaction value 
(₦937M). This gap is too large to reflect normal target-setting 
ambition and suggests the targets were generated using different 
assumptions, units, or scale than the transaction data.


## Conclusion 


This project shows a full fiancia analysis using excel then 
analyzing branch performance, channel reliability, customer 
segments, and account activity through PivotTables and a dashboard.

Key findings included a significant branch performance gap, a large 
share of inactive accounts, a disproportionately valuable High Risk 
segment, and a monthly targets file that doesn't reconcile with 
actual transactions . A reminder that numbers need questioning before 
they're used to judge performance.