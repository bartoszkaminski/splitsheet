function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Splitwise')
      .addItem('Update','updateExpenses')
      .addToUi();
}
 
function updateExpenses() {
   var service = getSplitwiseService();
   
   if (service.hasAccess()) {
     Logger.log("App has access.");
     
     var categories = getCategories();
     var currentUserId = getCurrentUserId();
     var expenses = getExpenses();
     var filteredExpenses = filterExpenses(expenses, currentUserId, categories);
     var sortedExpenses = sortExpenses(filteredExpenses);
     exportExpenses(sortedExpenses);
   }
   else {
     Logger.log("App has no access yet.");
     
     // open this url to gain authorization from Splitwise
     var authorizationUrl = service.getAuthorizationUrl();
     Logger.log("Open the following URL and re-run the script: %s",
         authorizationUrl);
   }
 }

function getExpenses() {
  var expensesPath = "https://secure.splitwise.com/api/v3.0/get_expenses";
  var headers = {
    "Authorization": "OAuth " + getSplitwiseService().getAccessToken()
  };
  
  var sheet = SpreadsheetApp.getActiveSheet();
  var month = sheet.getIndex();
  var from = new Date(2017,month-1,1); //months are 0-11
  var to = new Date(2017,month,1);
  var payload = {
    "limit": "500",
    "dated_after": from.toJSON(),
    "dated_before": to.toJSON()
  };
  
  var options = {
    "headers": headers,
    "payload": payload,
    "method" : "GET",
    "muteHttpExceptions": true
  };
  
  var expensesResponse = UrlFetchApp.fetch(expensesPath, options);
  var expenses = JSON.parse(expensesResponse.getContentText()).expenses;
  return expenses;
}

function filterExpenses(expenses, currentUserId, categories) {
  var expensesToReturn = [];
  for (i = 0; i < expenses.length; i++) {
    var fullExpense = expenses[i];
    if (fullExpense.deleted_at != null || fullExpense.category.id == 18) { continue; }
   
    var users = fullExpense.users;
    var cost = null;
    for (j = 0; j < users.length; j++) {
      if (users[j].user.id == currentUserId) {
        cost = users[j].owed_share;
      }
    }
    if (cost == null || cost == 0) { continue; }
  
    var expense = {
      date: new Date(fullExpense.date),
      description: fullExpense.description,
      category: categories[fullExpense.category.id].category,
      subcategory: categories[fullExpense.category.id].subcategory,
      cost: cost,
      currency: fullExpense.currency_code
    };
    expensesToReturn.push(expense);
  }
  return expensesToReturn;
}

function sortExpenses(expenses) {
  return expenses.sort(function(a,b) { return new Date(a.date) - new Date(b.date); });
}

function exportExpenses(expenses) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var allCells = sheet.getRange(3, 1, 197, 5);
  allCells.clearContent();
  allCells.setBackground("white");
  var firstCell = 3;
  for (i = 0; i < expenses.length; i++) {
    var expense = expenses[i];
    sheet.getRange(firstCell+i, 1).setValue(expense.date);
    sheet.getRange(firstCell+i, 2).setValue(expense.category);
    sheet.getRange(firstCell+i, 3).setValue(expense.subcategory);
    sheet.getRange(firstCell+i, 4).setValue(expense.description);
    if (expense.currency == "PLN") {
      sheet.getRange(firstCell+i, 5).setValue(expense.cost.replace(".", ","));
    } else {
      sheet.getRange(firstCell+i, 5).setValue(0);
      sheet.getRange(firstCell+i, 5).setBackground("red");
    }
  }
}

function getCurrentUserId() {
  var currentUserPath = "https://secure.splitwise.com/api/v3.0/get_current_user";
  var headers = {
       "Authorization": "OAuth " + getSplitwiseService().getAccessToken()
     };
     
     var options = {
       "headers": headers,
       "method" : "GET",
       "muteHttpExceptions": true
    };
 
  var userResponse = UrlFetchApp.fetch(currentUserPath, options);
  var currentUserId = JSON.parse(userResponse.getContentText()).user.id;
  return currentUserId;
}

function getCategories() {
  var categoriesPath = "https://secure.splitwise.com/api/v3.0/get_categories";
  var headers = {
       "Authorization": "OAuth " + getSplitwiseService().getAccessToken()
     };
     
     var options = {
       "headers": headers,
       "method" : "GET",
       "muteHttpExceptions": true
    };
  
  var categoriesResponse = UrlFetchApp.fetch(categoriesPath, options);
  var categories = JSON.parse(categoriesResponse.getContentText()).categories;
 
  var categoriesToReturn = [];
  for (i = 0; i < categories.length; i++) {
    var subcategories = categories[i].subcategories;
    for (j = 0; j < subcategories.length; j++) {
      var subcategory = {
        category: categories[i].name,
        subcategory: subcategories[j].name
      };
      categoriesToReturn[subcategories[j].id] = subcategory;
    }
  }
  return categoriesToReturn;
}
