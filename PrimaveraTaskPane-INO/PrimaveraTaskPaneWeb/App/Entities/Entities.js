﻿

function AppContext() {
    var name = "";
    
    return {
        getName: function (newName) {
            return name;
        },
        setName: function (newName) {
            name = newName;
        }
    }   
}

function Formula() {
    var associativeArray = {};
    var name = "";
    var cell = "A1";

    return {
        addParameter: function (key, value) {
            associativeArray[key] = value;
        },
        getParameters: function () {
            return associativeArray;
        },

        getParameter: function (key) {
            return associativeArray[key];
        },
        getName: function () {
            return name;
        },
        setName: function (newName) {
            name = newName;
        },
        getCell: function () {
            return cell;
        },
        setCell: function (newName) {
            cell = newName;
        }
    }
}

function ListFormulas() {
    var list = {};
  
    return {
        add: function (key, value) {
            list[key] = value;
        },
        getLists: function () {
            return list;
        },
        getList: function (key) {
            return list[key];
        }
    }
}