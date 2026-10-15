# Gank Sales Helper
This script generates the following files:
* `Self.cache` - Has your gank user id
* `Listings.cache` - Has all of your Gank listings.
* `Transactions.cache` Has all of your Gank transactions.
* `GankSales-YYYY-MM-DD-HH-MM-SS.csv` - The most recent sales report.

When this script runs, it gets your latest listings and transactions up until it finds one that's already in the cache file. This means it will take a very long time to run the first time, and then subsequent runs will be much faster.
This script requires you to paste your Gank bearer token. You can find this by going to the Gank website in a desktop web browser, hitting F12, going to the "Network" tab, clicking a link in Gank, finding a request in the Network tab with type `xhr`, going to its `Header` values, and copying the value for `Authorization`. Hint: It should look like `Bearer: eyJ...`

The output is a .csv file that can be opened by Excel or any Excel clone.
