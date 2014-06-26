
function loadLists(value){
	if(value == 'PRIBSS'){
		 return {
			    "company": "PRIBSS",
			   
			    "lists": [
			        {
			            "name": "Sales",
						"description" : "Sales"
			        },
			        {
			             "name": "Order",
						 "description" : "Order"
			        }
			    ]
			}
	}
	else {
		return {
			    "company": "PRITEC",
			   
			    "lists": [
			        {
			            "name": "Sales",
						"description" : "Sales Tec"
			        },
			        {
			             "name": "Order",
						 "description" : "Order Tec"
			        }
			    ]
			}
	}
}
	
function loadCompanies(){
  return {
    "organization": "Primavera",
   
    "companies": [
        {
            "name": "PRIBSS",
			"description" : "Primavera Bss"
        },
        {
             "name": "PRITEC",
			"description" : "Primavera Tec"
        }
    ]}
}
	
function loadOrderList(){
  return {
	  "data": [
		   {
            "key": "ENC.2014.1",
			"documentType" : "ENC",
			"serie" : "2014",
			"number" : "1",
			"supplier" : "SOFRIO",
			"total" : 212
			},
			 {
            "key": "ENC.2014.2",
			"documentType" : "ENC",
			"serie" : "2014",
			"number" : "2",
			"supplier" : "SOFRIO",
			"total" : 121
			}, {
            "key": "ENC.2014.3",
			"documentType" : "ENC",
			"serie" : "2014",
			"number" : "3",
			"supplier" : "SOFRIO",
			"total" : 242
			}
      ]
  }
}

	
function loadSalesList(){
  return {
	  "data": [
		   {
            "key": "FA.2014.1",
			"documentType" : "FA",
			"serie" : "2014",
			"number" : "1",
			"supplier" : "ALCAD",
			"total" : 212
			},
			 {
            "key": "FA.2014.2",
			"documentType" : "FA",
			"serie" : "2014",
			"number" : "2",
			"supplier" : "SOFRIO",
			"total" : 121
			}, {
            "key": "FA.2014.3",
			"documentType" : "FA",
			"serie" : "2014",
			"number" : "3",
			"supplier" : "ALCAD",
			"total" : 242
			}
      ]
  }
}