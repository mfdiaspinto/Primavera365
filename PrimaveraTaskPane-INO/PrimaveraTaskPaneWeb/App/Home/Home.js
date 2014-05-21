﻿/// <reference path="Templates/login.html" />
/// <reference path="Templates/Login.html" />
/// <reference path../../Scripts/App.js" />

(function () {
    "use strict";
    var server = 'http://priserver-mfdiaspinto.rhcloud.com';

    function getTextFromDocument() {

        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            { valueFormat: "unformatted", filterType: "all" },

            function (asyncResult) {
                showStockData(asyncResult.value);
            });

    }

    function addListToExcel() {

        var company = $('#companySelector').val();
        var list = $('#listSelector').val();

        var urlPath = server + '/' + company + '/' + list;
        $.ajax({
            type: 'GET',
            url: urlPath,
            success: function (data) {
                writeTableToExcel(data.columns, data.rows);
            },
            error: function (error) {
                write(error.statusText);
            }
        });
    }

    function addListToPanel() {

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

    function writeTableToExcel(columns, rows) {
        var myTable = new Office.TableData();

        var lineHeader = [];
        for (var i = 0; i < columns.length; i++) {

            lineHeader.push(columns[i].name);
        }
        // SET HEADER
        myTable.headers = [lineHeader];

        //SET ROWS
        for (var i = 0; i < rows.length; i++) {
            var line = [];

            for (var j = 0; j < columns.length; j++) {
                line.push(rows[i][columns[j].name]);
            }
            myTable.rows.push(line);
        }

        // Write table.
        Office.context.document.setSelectedDataAsync(myTable, { coercionType: "table" },
            function (result) {
                var error = result.error
                if (result.status === "failed") {
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

            var urlPath = server + '/companies';
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
        });
    }

    function loadListsForm() {
        $.get("Templates/lists.htm", '', function (data) {
            $("#listsDiv").append(data);
            $('#add-panel').click(addListToPanel);
            $('#add-excel').click(addListToExcel);

//            var e = document.getElementById("companySelector");
//            var strUser = e.options[e.selectedIndex].value;

            var company = $('#companySelector').val();

            var urlPath = server + '/' + company + '/lists';
            $.ajax({
                type: 'GET',
                url: urlPath,
                success: function (data) {
                        $('#listSelector').remove();
                        $('#divListSelector').append('<select id="listSelector" class="form-control"> </select>');
                        for (var i = 0; i < data.rows.length; i++) {
                            $('#listSelector').append('<option value=' + data.rows[i].key + '>' + data.rows[i].description + '</option>');
                        }
                },
                error: function (error) {
                    write(error.statusText);
                }
            });
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

        $.getJSON(server + "/netsales/" + formula.getParameter("Company"), function (data) {
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