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
     var travelGroupsIds = getTravelGroupsIds();
     var currentUserId = getCurrentUserId();
     var expenses = getExpenses();
     var filteredExpenses = filterExpenses(expenses, currentUserId, categories, travelGroupsIds);
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
  var from = sheet.getRange(1, 21).getValue();
  var to = sheet.getRange(2, 21).getValue();
  try {
    from.setSeconds(from.getSeconds() - 1);
    to.setDate(to.getDate() + 1);
  } catch(e) {
    throw 'Please specify correct date range';
  }
  
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

function filterExpenses(expenses, currentUserId, categories, travelGroupsIds) {
  var expensesToReturn = [];
  for (i = 0; i < expenses.length; i++) {
    var fullExpense = expenses[i];
    if (fullExpense.deleted_at != null || fullExpense.payment == true || fullExpense.category.id == 18) { continue; }
   
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
    var travelAwareExpense = markAsTravelIfNeeded(fullExpense, expense, travelGroupsIds);
    expensesToReturn.push(travelAwareExpense);
  }
  return expensesToReturn;
}

function markAsTravelIfNeeded(fullExpense, expense, travelGroupsIds) {
  if (travelGroupsIds.indexOf(fullExpense.group_id) > -1) {
    expense.category = "Entertainment";
    expense.subcategory = "Travel";
  }
  return expense;
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

function getTravelGroupsIds() {
  var groupsPath = "https://secure.splitwise.com/api/v3.0/get_groups";
  var headers = {
       "Authorization": "OAuth " + getSplitwiseService().getAccessToken()
     };
     
     var options = {
       "headers": headers,
       "method" : "GET",
       "muteHttpExceptions": true
    };
  
  var groupsResponse = UrlFetchApp.fetch(groupsPath, options);
  var groups = JSON.parse(groupsResponse.getContentText()).groups;
  
  var travelGroupsIdsToReturn = [];
  for (i = 0; i < groups.length; i++) {
    var group = groups[i];
    Logger.log(group.group_type);
    if (group.group_type == "trip" || group.group_type == "travel") {
      travelGroupsIdsToReturn.push(group.id);
    }
  }
  return travelGroupsIdsToReturn;
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
