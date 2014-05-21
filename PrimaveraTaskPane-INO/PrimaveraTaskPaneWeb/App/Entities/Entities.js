

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
    var key = "";
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
        getKey: function () {
            return key;
        },
        setKey: function (newName) {
            key = newName;
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
    var count = 0;
    return {
        add: function (key, value) {
            if (list[key] == undefined)
            {
                list[key] = value;
                count = count + 1;
            }

            list[key] = value;
        },
        getLists: function () {
            return list;
        },
        getList: function (key) {
            return list[key];
        },
        getCount: function (key) {
           return count;
        }
    }
}