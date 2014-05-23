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
        var company = $('#companySelector').val();
        var list = $('#listSelector').val();

        var urlPath = server + '/' + company + '/' + list;
        $.ajax({
            type: 'GET',
            url: urlPath,
            success: function (data) {
                writeTableToPanel(data.columns, data.rows);
            },
            error: function (error) {
                write(error.statusText);
            }
        });
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

    function writeTableToPanel(columns, rows) {
        $('#idTable').remove();
        $('#box').append('<table id="idTable" class="table table-condensed"></table>');
        $('#idTable').append('<thead id="idColumn"> </thead>');
        $('#idColumn').append('<tr id="idTr"> </tr>');
        for (var i = 0; i < columns.length; i++) {
            $('#idTr').append('<th>' + columns[i].name + '</th>');
        }
        $('#idTable').append('<tbody id="idTbody"> </thead>');

        for (var i = 0; i < rows.length; i++) {

            var idrow = 'idRow' + i;

            $('#idTbody').append('<tr id="' + idrow + '"></tr>');
            for (var j = 0; j < columns.length; j++) {

                $('#' + idrow).append('<td>' + rows[i][columns[j].name] + '</td>');
            }
        }
    }

    function write(message) {
        $('#errorModal').modal('show');
        document.getElementById('message').innerText = "";
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
        var formulaCount = listFormulas.getCount();
        formula.setKey(formulaCount + '_NetSales');
        formula.setName('NetSales_' + $('#newFormulaCompany').val() + '_' + $('#newFormulaYear').val());
        formula.setCell($('#newFormulaCellID').val());
        formula.addParameter("Company", $('#newFormulaCompany').val());
        formula.addParameter("Year", $('#newFormulaYear').val());
        formula.addParameter("Month", $('#newFormulaMonth').val());

        listFormulas.add(formula.getKey(), formula);
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
                                else
                                {
                                    $('#formulasDiv').append('<div class="row"><a id=' + formula.getKey() + ' href="#">' + formula.getName() + '</a></div>');
                                    $('#' + formula.getKey()).click({ param1: formula.getKey() }, openFormula);
                                    $('#myModal').modal('hide');
                                    cleanNewFormula();
                                }
                            });
                    }
                });
            })
          .fail(function (data) {
              console.log("error");
          });
    }

    function editFormula() {
        var key = $('#editFormulaId').text();
        var formula = listFormulas.getList(key);

        formula.setCell($('#editFormulaCellID').val());
        formula.addParameter("Company", $('#editFormulaCompany').val());
        formula.addParameter("Year", $('#editFormulaYear').val());
        formula.addParameter("Month", $('#editFormulaMonth').val());

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
                                          formula.setName('NetSales_' + formula.getParameter("Company") + '_' + formula.getParameter("Year"));
                                          listFormulas.add(formula.getKey(), formula);
                                          $('#' + formula.getKey()).text(formula.getName());
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
        var key = event.data.param1;
        var formula = listFormulas.getList(key);
        $('#editFormulaId').text(key);
        $('#editFormulaName').text(formula.getName());
        $('#editFormulaYear').val(formula.getParameter("Year"));
        $('#editFormulaMonth').val(formula.getParameter("Month"));
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

    function refreshAllFormulas() {
        refreshFormula(0);
    }

    function refreshFormula(id) {
        var formula = listFormulas.getList(id + '_NetSales');
        if(formula != undefined) {
            $.getJSON(server + "/netsales/" + formula.getParameter("Company"), function (data) {

                Office.context.document.bindings.addFromNamedItemAsync(formula.getCell(), Office.BindingType.Text, { id: formula.getKey() },
                          function (asyncResult) {
                              if (asyncResult.status == "failed") {
                                  write('Error: ' + asyncResult.error.message);
                              }
                              else {
                                  // Write data to the new binding.
                                  Office.select("bindings#" + formula.getKey()).setDataAsync(data.value, { coercionType: Office.BindingType.Text },
                                      function (asyncResult) {
                                          if (asyncResult.status == "failed") {
                                              write('Error: ' + asyncResult.error.message);
                                          }
                                          else
                                          {
                                              refreshFormula(id +1);
                                          }
                                      });
                              }
                          });
            })
              .fail(function (data) {
                  write("Error in server");
              });
        }
    }

    function loadCompanyLists() {
        loadListsForm();
    }
    
    function generatePath(parameters) {

       
        return server + "";
    }


    // INIT APP
    Office.initialize = function (reason) {
            $(document).ready(function () {
		    $('#get-text').click(getTextFromDocument);
	   
		    appContext.setName("Miguel Dias");
		    var textWelcome = '<h5>Welcome ' + appContext.getName() + '</h5>';
		    $('#loggedinuser').append(textWelcome);

		    $('#addFormula').click(addNewFormula);
		    $('#refreshAll').click(refreshAllFormulas);

		    $('#cleanNewFormula').click(cleanNewFormula);

		    $('#editFormula').click(editFormula);
		    $('#cleanEditFormula').click(cleanEditFormula);

		    loadCompaniesForm();
            });
        }

    // VARIAVEIS GLOBAIS
    var appContext = new AppContext();
    var listFormulas = new ListFormulas();
    var server = 'http://priserver-mfdiaspinto.rhcloud.com';
})();