/// <reference path="Templates/login.html" />
/// <reference path="Templates/Login.html" />
/// <reference path../../Scripts/App.js" />

(function () {
    "use strict";

    function getTextFromDocument() {

        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            { valueFormat: "unformatted", filterType: "all" },

            function (asyncResult) {
                showStockData(asyncResult.value);
            });

    }

    function addListToExcel() {
	    var e = document.getElementById("listSelector");
	    var strUser = e.options[e.selectedIndex].value;
	    if("Order" == strUser)
	    {
	      var list = loadOrderList();
	      writeTableToExcel(list.data);
	    }
	    else
	    {
		    var list = loadSalesList();
		    writeTableToExcel(list.data);
	    }
    }

    function testApi() {

     var test = 'http://query.yahooapis.com/v1/public/yql?q=select%20*%20from%20answers.getbycategory%20where%20category_id%3D2115500137%20and%20type%3D%22resolved%22&format=json&diagnostics=true&callback=';
  
      $.getJSON( test, function(result) {
        console.log( "success" );
      })
        .done(function(result) {
            writeTableToExcel(result.query.results.Question);
        })
        .fail(function(error) {
          console.log( "error" );
        })
        .always(function() {
          console.log( "complete" );
        })
   
}

    function addListToPanel() {
    
      /* var test = 'http://query.yahooapis.com/v1/public/yql?q=select%20*%20from%20answers.getbycategory%20where%20category_id%3D2115500137%20and%20type%3D%22resolved%22&format=json&diagnostics=true&callback=';

	    $.getJSON( test, function(result) {
	      console.log( "success" );
	    })
	      .done(function(result) {
		    writeTableToExcel(result);
	      })
	      .fail(function(error) {
	        console.log( "error" );
	      })
	      .always(function() {
	        console.log( "complete" );
	      });
	      */
	    var e = document.getElementById("listSelector");
	    var strUser = e.options[e.selectedIndex].value;
	    if("Order" == strUser)
	    {
	      var list = loadOrderList();
	  
	      writeTableToPanel(list);
	    }
	    else
	    {
		     var list = loadSalesList();
		     writeTableToPanel(list);
	    }
    }
   
    function login() {
        //var postData = { username: "vitor.costa@primaverabss.com", password: "aaa" };
   
        var postData = { username: $('#userName').val(), password: $('#password').val() };

        $.ajax({
            type: "POST",
            //url: 'https://testsvc.cloudprimavera.com/mobile/mvchost/api/login',
            url: 'http://localhost:51929/api/login',
            // The key needs to match your method's input parameter (case-sensitive).
            data: postData,
            headers: {
                "Accept" : "application/json; charset=utf-8",
                'Access-Control-Allow-Origin' : '*'
            },
            dataType: "json",
            success: function (data) {
                var textWelcome = '<h5>Welcome ' + data.name + '</h5>'; 
                $('#loggedinuser').append(textWelcome);
                $('#loginDiv').remove();
                loadCompaniesForm();
            },
            error: function (error) {
                write(error.statusText);
            }
        });

    }

    function writeTableToExcel(result) {

        //Office.context.document.bindings.addFromNamedItemAsync("A10:A13", "matrix", { id: "MyCities" },
        //function (asyncResult) {
        //    if (asyncResult.status == "failed") {
        //        write('Error: ' + asyncResult.error.message);
        //    }
        //    else {
        //        // Write data to the new binding.
        //        Office.select("bindings#MyCities").setDataAsync([['Berlin'], ['Munich'], ['Duisburg']], { coercionType: "matrix" },
        //            function (asyncResult) {
        //                if (asyncResult.status == "failed") {
        //                    write('Error: ' + asyncResult.error.message);
        //                }
        //            });
        //    }
        //});

        $('#loggedinuser').append(appContext.getName);

	    var rows = result.length;
	    var myTable = new Office.TableData();
        myTable.headers = [["key", "documentType", "serie", "number", "supplier", "total"]];
	    myTable.rows = [];
	    var items = result;

	    for (var i = 0; i < rows; i++) {
		    myTable.rows.push([items[i].key, items[i].documentType, items[i].serie, items[i].number, items[i].supplier, items[i].total]);
	    }
	
        Office.context.document.setSelectedDataAsync(myTable, {coercionType: "table"},
            function (asyncResult) {
                var error = asyncResult.error;
                if (asyncResult.status === "failed"){
                write(error.name + ": " + error.message);
                }
            });
    }

    function writeTableToPanel(result) {
	    var rows = result.data.length;
	    var myTable = new Office.TableData();
        myTable.headers = [["key", "documentType", "serie", "number", "supplier", "total"]];
	    myTable.rows = [];
	    var items = result.data;
	    $('#idTable').remove();
	    $('#box').append('<table id="idTable" class="table table-condensed"></table>');
	    $('#idTable').append('<thead id="idColumn"> </thead>');
	    $('#idColumn').append('<tr id="idTr"> </tr>');
	    $('#idTr').append('<th>Key</th>');
	    $('#idTr').append('<th>Document Type</th>');
	    $('#idTr').append('<th>Serie</th>');
	    $('#idTr').append('<th>Number</th>');
	    $('#idTr').append('<th>Supplier</th>');
	    $('#idTr').append('<th>Total</th>');
	    $('#idTable').append('<tbody id="idTbody"> </thead>');

	    for (var i = 0; i < rows; i++) {
		    var idrow = 'idRow'+i;
		    $('#idTbody').append('<tr id="'+idrow + '"></tr>');
		    $('#' + idrow).append('<td>' +items[i].key + '</td>');
		    $('#' + idrow).append('<td>' +items[i].documentType + '</td>');
		    $('#' + idrow).append('<td>' + items[i].serie+ '</td>');
		    $('#' + idrow).append('<td>' + items[i].number+ '</td>');
		    $('#' + idrow).append('<td>' +items[i].supplier + '</td>');
		    $('#' + idrow).append('<td>' +items[i].total + '</td>');
	    }
    }

    // Function that writes to a div with id='message' on the page.
    function write(message){
        document.getElementById('message').innerText += message; 
    }

    function loadLoginForm() {
        $.get("Templates/login.htm", '', function (data) {
            $("#loginDiv").append(data);
            $('#login').click(login);
        });
    }

    function loadCompaniesForm() {
        $.get("Templates/companies.htm", '', function (data) {
            $("#companyDiv").append(data);
            $('#load-lists').click(loadCompanyLists);

            var urlPath = 'http://priserver-mfdiaspinto.rhcloud.com/companies';

            $.ajax({
                type: 'GET',
                url: urlPath,
                success: function (data) {
                  for (var i = 0; i < data.length; i++) {
                    $('#companySelector').append('<option value=' + data[i].name + '>' + data[i].name + '</option>');
                    $('#editFormulaCompany').append('<option value=' + data[i].name + '>' + data[i].name + '</option>');
                    $('#newFormulaCompany').append('<option value=' + data[i].name + '>' + data[i].name + '</option>');

                  }
                },
                error: function (error) {
                    write(error.statusText);
                }
            });

//            var data = loadCompanies();

           
        });
    }

    function loadListsForm() {
        $.get("Templates/lists.htm", '', function (data) {
            $("#listsDiv").append(data);
            $('#add-panel').click(addListToPanel);
            $('#add-excel').click(addListToExcel);
            $('#list-api').click(testApi);

            var e = document.getElementById("companySelector");
            var strUser = e.options[e.selectedIndex].value;

            var data = loadLists(strUser);
            $('#listSelector').remove();
            $('#divListSelector').append('<select id="listSelector" class="form-control"> </select>');
            for (var i = 0; i < data.lists.length; i++) {
                $('#listSelector').append('<option value=' + data.lists[i].name + '>' + data.lists[i].description + '</option>');
            }
        });
    }

    function addNewFormula() {
        var formula = new Formula();
        formula.setName($('#newFormulaName').val());
        formula.setCell($('#newFormulaCellID').val());
        formula.addParameter("Company", $('#newFormulaCompany').val());
        formula.addParameter("Year", $('#newFormulaYear').val());

        listFormulas.add(formula.getName(), formula);
       
        Office.context.document.bindings.addFromNamedItemAsync(formula.getCell(), Office.BindingType.Text, { id: "PriFormula" },
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    write('Error: ' + asyncResult.error.message);
                }
                else {
                    // Write data to the new binding.
                    Office.select("bindings#PriFormula").setDataAsync("Result Formula" + formula.getCell(), { coercionType: Office.BindingType.Text },
                        function (asyncResult) {
                            if (asyncResult.status == "failed") {
                                write('Error: ' + asyncResult.error.message);
                            }
                            else
                            {
                                $('#formulasDiv').append('<div class="row"><a id=' + formula.getName() + ' href="#">' + formula.getName() + '</a></div>');
                                $('#' + formula.getName()).click({ param1: formula.getName() }, openFormula);
                                $('#myModal').modal('hide');
                                cleanNewFormula();
                            }
                        });
                }
            });
    }

    function editFormula() {
        var name = $('#editFormulaName').text();
        var formula = listFormulas.getList(name);

        formula.setCell($('#editFormulaCellID').val());
        formula.addParameter("Company", $('#editFormulaCompany').val());
        formula.addParameter("Year", $('#editFormulaYear').val());

        $.getJSON("http://priserver-mfdiaspinto.rhcloud.com/netsales/" + formula.getParameter("Company"), function (data) {
            Office.context.document.bindings.addFromNamedItemAsync(formula.getCell(), Office.BindingType.Text, { id: "PriFormula" },
                      function (asyncResult) {
                          if (asyncResult.status == "failed") {
                              write('Error: ' + asyncResult.error.message);
                          }
                          else {
                              // Write data to the new binding.
                              Office.select("bindings#PriFormula").setDataAsync(data.value, { coercionType: Office.BindingType.Text },
                                  function (asyncResult) {
                                      if (asyncResult.status == "failed") {
                                          write('Error: ' + asyncResult.error.message);
                                      }
                                      else {
                                          listFormulas.add(formula.getName(), formula);
                                          cleanEditFormula();
                                      }
                                  });
                          }
                      });
        })
          .fail(function (data) {
              console.log("error");
          });
    }

    function openFormula(event) {
        var name = event.data.param1;
        var formula = listFormulas.getList(name);
        $('#editFormulaName').text(formula.getName());
        $('#editFormulaYear').val(formula.getParameter("Year"));
        $('#editFormulaCompany').val(formula.getParameter("Company"));
        $('#editFormulaCellID').val(formula.getCell());

        $('#modelFormulaEdit').modal('show');
    }

    function cleanEditFormula() {
        $('#editFormulaName').text("");
        $('#editFormulaYear').val("");
        $('#editFormulaCompany').val("");
        $('#editFormulaCellID').val("");
        $('#modelFormulaEdit').modal('hide');
    }

    function cleanNewFormula() {
        $('#newFormulaName').val("");
        $('#newFormulaYear').val("");
        $('#newFormulaCompany').val("");
        $('#newFormulaCellID').val("");
        $('#modelFormulaEdit').modal('hide');
    }

    Office.initialize = function (reason) {
            $(document).ready(function () {
		    $('#get-text').click(getTextFromDocument);
	   
		    appContext.setName("Miguel Dias");
		    var textWelcome = '<h5>Welcome ' + appContext.getName() + '</h5>';
		    $('#loggedinuser').append(textWelcome);

		    $('#addFormula').click(addNewFormula);
		    $('#cleanNewFormula').click(cleanNewFormula);

		    $('#editFormula').click(editFormula);
		    $('#cleanEditFormula').click(cleanEditFormula);

		    loadCompaniesForm();
            });
        }

    function loadCompanyLists(){
        loadListsForm();
    }

    var appContext = new AppContext();
    var listFormulas = new ListFormulas();
})();