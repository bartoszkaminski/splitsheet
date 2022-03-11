function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Splitwise')
      .addItem('Update','updateExpenses')
      .addToUi();
}
 
function updateExpenses() {
   var service = getSplitwiseService();
   
   if (service.hasAccess()) {     
     var categories = getCategories();
     var tripGroupsIds = getTripGroupsIds();
     var currentUserId = getCurrentUserId();
     var expenses = getExpenses();
     var filteredExpenses = filterExpenses(expenses, currentUserId, categories, tripGroupsIds);
     var sortedExpenses = sortExpenses(filteredExpenses);
     exportExpenses(sortedExpenses);
   }
   else {     
     var authorizationUrl = service.getAuthorizationUrl();
     var ui = SpreadsheetApp.getUi();
     ui.alert("Spreadsheet has no access yet. Open the following URL to authorize this spreadsheet in Splitwise and try again: " + authorizationUrl);
   }
 }

function filterExpenses(expenses, currentUserId, categories, tripGroupsIds) {
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
    var tripAwareExpense = markAsTripIfNeeded(fullExpense, expense, tripGroupsIds);
    expensesToReturn.push(tripAwareExpense);
  }
  return expensesToReturn;
}

function markAsTripIfNeeded(fullExpense, expense, tripGroupsIds) {
  if (tripGroupsIds.indexOf(fullExpense.group_id) > -1) {
    expense.category = "Entertainment";
    expense.subcategory = "Trips";
  }
  return expense;
}

function sortExpenses(expenses) {
  return expenses.sort(function(a,b) { return new Date(a.date) - new Date(b.date); });
}

function exportExpenses(expenses) {
  var configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  var userCurrency = configSheet.getRange(2, 4).getValue();
  var sheet = SpreadsheetApp.getActiveSheet();
  var allCells = sheet.getRange(3, 1, 197, 5);
  allCells.clearContent();
  allCells.setBackground("white");
  var firstCell = 3;
  for (i = 0; i < expenses.length; i++) {
    var expense = expenses[i];
    var cost = expense.cost.replace(".", ",");
    sheet.getRange(firstCell+i, 1).setValue(expense.date);
    sheet.getRange(firstCell+i, 2).setValue(expense.category);
    sheet.getRange(firstCell+i, 3).setValue(expense.subcategory);
    sheet.getRange(firstCell+i, 4).setValue(expense.description);
    if (expense.currency == userCurrency) {
      sheet.getRange(firstCell+i, 5).setValue(cost);
    } else {
      sheet.getRange(firstCell+i, 5).setNote(cost + " " + expense.currency)
      sheet.getRange(firstCell+i, 5).setFormula('=Index(GOOGLEFINANCE("CURRENCY:' + expense.currency + userCurrency + '";"price";A' + (firstCell+i) + ';0;"DAILY");2;2)*' + cost)
    }
  }
}

// Splitwise API
function getExpenses() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const from = sheet.getRange(1, 21).getValue();
  const to = sheet.getRange(2, 21).getValue();
  try {
    from.setSeconds(from.getSeconds() - 1);
    to.setDate(to.getDate() + 1);
  } catch(e) {
    throw 'Please specify correct date range';
  }
  
  const expensesPath = "https://secure.splitwise.com/api/v3.0/get_expenses?limit=500&dated_after="+from.toJSON()+"&dated_before="+to.toJSON();
  const expensesResponse = callSplitwiseAPI(expensesPath);
  return expensesResponse.expenses;
}

function getCurrentUserId() {
  const currentUserPath = "https://secure.splitwise.com/api/v3.0/get_current_user";
  const userResponse = callSplitwiseAPI(currentUserPath);
  return userResponse.user.id;
}

function getTripGroupsIds() {
  const groupsPath = "https://secure.splitwise.com/api/v3.0/get_groups";
  const groupsResponse = callSplitwiseAPI(groupsPath);
  
  var tripGroupsIdsToReturn = [];
  for (const group of groupsResponse.groups) {
    if (group.group_type == "trip" || group.group_type == "travel") {
      tripGroupsIdsToReturn.push(group.id);
    }
  }
  return tripGroupsIdsToReturn;
}

function getCategories() {
  const categoriesPath = "https://secure.splitwise.com/api/v3.0/get_categories"; 
  const categoriesResponse = callSplitwiseAPI(categoriesPath);
 
  const categoriesToReturn = [];
  for (const cat of categoriesResponse.categories) {
    for (const subcat of cat.subcategories) {
      categoriesToReturn[subcat.id] = {
        category: cat.name,
        subcategory: subcat.name
      };
    }
  }
  return categoriesToReturn;
}

function callSplitwiseAPI(url, options={}) {
  options.headers = Object.assign({
    Authorization: "OAuth " + getSplitwiseService().getAccessToken(),
  }, options.headers);
  response = UrlFetchApp.fetch(url, options);
  return JSON.parse(response.getContentText());
}
