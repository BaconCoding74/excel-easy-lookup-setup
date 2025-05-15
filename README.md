# Excel Lookup Setup for Excel New Users
This project is inspired by complaints of teacher in manual grade entering in Excel when they did not learn how to use Excel formula before. 
This program allows the user to setup lookup formula by using named range. Hence, new user can also exposed to named range and lookup formula which is quite useful in automating repetitive task.
This macro allows user to setup lookup formula with easy learning curve.

## Lookup Function
```LOOKUP(lookup_value, lookup_vector, [result_vector])```
lookup_value: value used to searching in the vector range provided
lookup_vector: range of value used to compare with lookup_value
result_vector: range of value displayed based on final comparison result from lookup

## Simple Sample about Range Required for Lookup Function in Excel
![image](https://github.com/user-attachments/assets/4deb63b6-125c-4e2a-b917-79ff0e4056e7)
Input range: range of input (lookup_value)
Output range: range of output (result of lookup)
Vector range: range of value for conditional comparison (lookup_vector)
Result range: range of value to be displayed (result_vector)

## UI For Lookup Setup
![image](https://github.com/user-attachments/assets/b8d5b551-db2d-4516-b99a-daa14bedad7d)
### Create named range
It allows new user to create named range after selecting the area and type the name for the range.
### Find named range
It allows users to search named range for validation purposes.
### Lookup function setup
It allows users to setup the lookup formula after defining four range required
