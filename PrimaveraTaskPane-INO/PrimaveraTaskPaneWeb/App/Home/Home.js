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

    function loadFormulas() {
       
            var urlPath = server + '/formulas';
            $.ajax({
                type: 'GET',
                url: urlPath,
                success: function (data) {
                    formulasTemplate = data;
                },
                error: function (error) {
                    write(error.statusText);
                }
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
        formula.setKey(formulaCount + '_FORMULA');
        formula.setCell($('#newFormulaCellID').val());
        formula.setFormulaName($('#newFormula').val());

        formula.addParameter("Company", $('#newFormulaCompany').val());
        formula.addParameter("Year", $('#newFormulaYear').val());
        formula.addParameter("Month", $('#newFormulaMonth').val());
        formula.addParameter("Day", "undefined");
        formula.setName(formula.getFormulaName() + '/' + formula.getParameter("Company") + '/' + formula.getParameter("Year") + '/' + formula.getParameter("Month"));

        $.getJSON(server + formula.getParameter("Company") + "/netsales/" + formula.getParameter("Year") + "/" + formula.getParameter("Month") + "/" + formula.getParameter("Day"), function (data) {
            Office.context.document.bindings.addFromNamedItemAsync(formula.getCell(), Office.BindingType.Text, { id: "PriFormula" },
                function (asyncResult) {
                    if (asyncResult.status == "failed") {
                        write('Error: ' + asyncResult.error.message);
                    }
                    else {
                        // Write data to the new binding.
                        if (data.length > 0) {
                            var result = data[0].value;
                            listFormulas.add(formula.getKey(), formula);

                            Office.select("bindings#PriFormula").setDataAsync(result, { coercionType: Office.BindingType.Text },
                                function (asyncResult) {
                                    if (asyncResult.status == "failed") {
                                        write('Error: ' + asyncResult.error.message);
                                    }
                                    else {
                                        $('#formulasDiv').append('<div class="row"><a id=' + formula.getKey() + ' href="#">' + formula.getName() + '</a></div>');
                                        $('#' + formula.getKey()).click({ param1: formula.getKey() }, openFormula);
                                        $('#myModal').modal('hide');
                                        cleanNewFormula();
                                    }
                                });
                        }
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
        formula.addParameter("Day", "undefined");


        $.getJSON(server + formula.getParameter("Company") + "/netsales/" + formula.getParameter("Year") + "/" + formula.getParameter("Month") + "/" + formula.getParameter("Day"), function (data) {
            Office.context.document.bindings.addFromNamedItemAsync(formula.getCell(), Office.BindingType.Text, { id: "PriFormula" },
                      function (asyncResult) {
                          if (asyncResult.status == "failed") {
                              write('Error: ' + asyncResult.error.message);
                          }
                          else {

                              if (data.length > 0) {
                                  var result = data[0].value;

                                  // Write data to the new binding.
                                  Office.select("bindings#PriFormula").setDataAsync(result, { coercionType: Office.BindingType.Text },
                                      function (asyncResult) {
                                          if (asyncResult.status == "failed") {
                                              write('Error: ' + asyncResult.error.message);
                                          }
                                          else {
                                              formula.setName(formula.getFormulaName() + '/' + formula.getParameter("Company") + '/' + formula.getParameter("Year") + '/' + formula.getParameter("Month"));
                                              listFormulas.add(formula.getKey(), formula);
                                              $('#' + formula.getKey()).text(formula.getName());
                                              cleanEditFormula();
                                          }
                                      });
                              }
                              else
                              {
                                  write("Error: empty.");
                              }
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
//        $('#editFormulaYear').val(formula.getParameter("Year"));
//        $('#editFormulaMonth').val(formula.getParameter("Month"));
        $('#editFormulaCompany').val(formula.getParameter("Company"));
        $('#editFormulaCellID').val(formula.getCell());

        $('#editParametersModal').empty();
        
        for (var i = 0; i < formulasTemplate[0].parameters.length; i++) {
            if (formulasTemplate[0].parameters[i].name != "Company" && formulasTemplate[0].parameters[i].name != "Day") {
                var nameValue = formulasTemplate[0].parameters[i].name;
                $('#editParametersModal').append('<div class="form-group"><label >' + nameValue + '</label><input id="editFormula' + nameValue + '" type="' + formulasTemplate[0].parameters[i].type + '" class="form-control" placeholder="Value" value="' + formula.getParameter(nameValue) + '"></div>');
            }
        }

        $('#modelFormulaEdit').modal('show');
    }

    function cleanEditFormula() {
        $('#editFormulaName').text("");
        $('#editFormulaYear').val("");
        $('#editFormulaCompany').val("");
        $('#editFormulaMonth').val("");

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
        var formula = listFormulas.getList(id + '_FORMULA');
        if (formula != undefined) {
            $.getJSON(server + formula.getParameter("Company") + "/netsales/" + formula.getParameter("Year") + "/" + formula.getParameter("Month") + "/" + formula.getParameter("Day"), function (data) {

                Office.context.document.bindings.addFromNamedItemAsync(formula.getCell(), Office.BindingType.Text, { id: formula.getKey() },
                          function (asyncResult) {
                              if (asyncResult.status == "failed") {
                                  write('Error: ' + asyncResult.error.message);
                              }
                              else {
                                  if (data.length > 0) {
                                      var result = data[0].value;

                                      // Write data to the new binding.
                                      Office.select("bindings#" + formula.getKey()).setDataAsync(result, { coercionType: Office.BindingType.Text },
                                          function (asyncResult) {
                                              if (asyncResult.status == "failed") {
                                                  write('Error: ' + asyncResult.error.message);
                                              }
                                              else {
                                                  refreshFormula(id + 1);
                                              }
                                          });
                                  }
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

    function generateNewFormulaModal() {

        $('#newParametersModal').empty();
        $('#newFormula').empty();
        for (var i = 0; i < formulasTemplate.length; i++) {
            $('#newFormula').append('<option value=' + formulasTemplate[i].name + '>' + formulasTemplate[i].name + '</option>');
        }

        for (var i = 0; i < formulasTemplate[0].parameters.length; i++) {
            if (formulasTemplate[0].parameters[i].name != "Company" && formulasTemplate[0].parameters[i].name != "Day") {
                $('#newParametersModal').append('<div class="form-group"><label >' + formulasTemplate[0].parameters[i].name + '</label><input id="newFormula' + formulasTemplate[0].parameters[i].name + '" type="' + formulasTemplate[0].parameters[i].type + '" class="form-control" placeholder="Value"></div>');
            }
        }

        $('#myModal').modal('show');
    }

    function saveReport() {
        var reportName = $('#newReportName').val();
        var list = listFormulas.getLists();
        var report = [];
        for (var i = 0; i < listFormulas.getCount() ; i++) {
            var idformula = i + '_FORMULA';
            report.push({
                key: list[idformula].getKey(),
                name: list[idformula].getName(),
                formulaName: list[idformula].getFormulaName(),
                cell: list[idformula].getCell(),
                parameters: list[idformula].getParameters(),
                });
        }

        var urlPath = server + 'report';
        $.ajax({
                type: 'POST',
                url: urlPath,
                data: { name: reportName, formulas: report },
                success: function (data) {
                    $('#myReportSave').modal('hide');
                },
                error: function (error) {
                    write(error.statusText);
                }
            });
    }


    function getReports() {
        $.getJSON(server + 'reports', function (data) {
            listFormulas = new ListFormulas();
            for (var i = 0; i < data.length ; i++) {
                var report = data[i].value;
                for (var j = 0; j < report.length ; j++) {
                    var formula = new Formula();
                    formula.setKey(report[j].key);
                    formula.setCell(report[j].cell);
                    formula.setFormulaName(report[j].formulaName);;
                    formula.setName(report[j].name);
                    formula.addParameter("Company", report[j].parameters.Company);
                    formula.addParameter("Year", report[j].parameters.Year);
                    formula.addParameter("Month", report[j].parameters.Month);
                    formula.addParameter("Day", "undefined");
                }
            }
        })
          .fail(function (data) {
              console.log("error");
          });
    }


    // INIT APP
    Office.initialize = function (reason) {
            $(document).ready(function () {
		    $('#get-text').click(getTextFromDocument);
	   
		    appContext.setName("Miguel Dias");
		    var textWelcome = '<h5>Welcome ' + appContext.getName() + '</h5>';
		    $('#loggedinuser').append(textWelcome);

		    $('#createFormula').click(generateNewFormulaModal);

		    $('#addFormula').click(addNewFormula);
		    $('#refreshAll').click(refreshAllFormulas);

		    $('#cleanNewFormula').click(cleanNewFormula);

		    $('#editFormula').click(editFormula);
		    $('#cleanEditFormula').click(cleanEditFormula);
		    $('#saveReport').click(saveReport);
		    
		    loadCompaniesForm();

		    loadFormulas();

		    getReports();
            });
        }

    // VARIAVEIS GLOBAIS
    var appContext = new AppContext();
    var listFormulas = new ListFormulas();
    var formulasTemplate;
    var server = 'http://priserver-mfdiaspinto.rhcloud.com/';
})();