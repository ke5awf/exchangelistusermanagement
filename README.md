# exchangelistusermanagement
Script to sync lists of users from SQL and add AD users to exchange distribution group. 


This is a script I made in an effort to eliminate an old mailman server. 

The Idea is to run this script at some regular interval, in my case nightly to take a list of users based on an SQL query, add those users to an exchange distribution group. 


This script assumes makes the assumption you have the rights to make these modifications to the exchange server via the powershell. 


If you have any suggestions to modifications feel free to post them up. 

Included is also an example Task for the task scheduler, that needs to be modified with User_ID, User_ID_string, and location of file. 
Keith
