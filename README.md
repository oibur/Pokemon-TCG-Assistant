This program will give you an excel sheet with the all standard Pokemon cards released to date and current prices for each. If you import a csv of your collection, it will also list all doubles you have. I am still working on matching the price sheet with collection csv to give a value of your current collection. I am also working on generating a list of cards you're missing.









March 25th, 2024
I've got a lot done since the last check-in: I figured out how to pull prices for all six formats instead of just unlimited. I was able to combine three files into one script instead of having one to pull set names, one to pull cards using set excel and one to pull prices from excel sheet. I was able to figure out why pricing my own cards wasn't working and need to fix that. 
Updated list of goals:
From JAN 2:
    -What is the value of my collection
    -What are the cards I am missing
    -Biggest price changes for cards each week
From JAN 26:
    -What is the value of my collection (broken down)
From FEB 2:
    -Get my cards, doubles and missing to generate to a single excel

February 2nd, 2024
After figuring out I wasn't working with a master list I figured out how to get that done.I figured out how to first generate a list of all sets and then pull cards up one set at a time and write them to an excel. Next I will:
1. Add my cards from csv to a new sheet in Excel
2. Get write doubles to write to a new sheet in the same Excel
3. Get write missing to write to a new sheet in the same Excel
4. Get everything to just create a new Excel everytime instead of updating one
    -this will not only simplify the code, but will also give data over time to analyze.
5. Combine files so everything can be done at once

January 26th, 2024
I have written scripts to: 1. write my doubles 2. write my missing 3. update a master excel in prices folder that will keep track of prices over time. I am still trying to figure out how to 
1. pull prices of unlimited AND reverse 
2. multiply price by column quantity AND add unlimited and reverse totals together if I have both formats so that 
3. One total price is added per row (under the column title being the date) even if there are multiple columns for that row to be calculated.
4. Continue with rest of requirements outlined below


January 22nd, 2024
My aim is to create an app that can help me keep track of all my Pokemon cards. 
1. What doubles do I have available to trade
    -Able to filter by set, reverse holo, rarity, etc. 
2. What is the value of my entire collection
3. What is the value of individual cards
4. What cards have changed a lot in value lately
    -Top 5 weekly drops and top 5 weekly jumps
5. What are the cheapest cards to buy that I don't have yet

I will start off by tackling each of these features individually, but hope to tie it all together into a single app with interactive menus.