// Objective Test

let HeaderObj = {}
let shxDept = ['SWE','ADX','APDM','CCS']
let nnHeaders = ["Year","Name","Joining","Dept"]
let nnDataArray = [['2021','Anurag','01/18/2022','SWE'],['','Shipra','09/18/2022','ADX']]
let hiringDataArray = []

nnHeaders.forEach((item, ind) => {
    HeaderObj[item] = ind
})

nnDataArray.forEach(rec => {
    nnHeaders.forEach((heading, ind) => {
        HeaderObj[heading] = rec[ind]
    })
    //console.log(HeaderObj)
    let targetYear = HeaderObj['Year']
    let candName = HeaderObj['Name']
    let startDate = HeaderObj['Joining']
    let entity = HeaderObj['Dept']

    //console.log(changeEntity(entity))
    hiringDataArray.push([changeYear(targetYear),candName,startDate,changeEntity(entity)])
})

function changeEntity(entity) {
    shxDept.forEach(dept => {
        if (entity == dept) {
            entity = 'SHX'
        }
    })
    return entity
}

function changeYear(targetYear) {
    if (targetYear == '') {
        targetYear = '2023'
    }
    return targetYear
}

function ExcelDateToJSDate(serial) {
    var utc_days  = Math.floor(serial - 25569);
    var utc_value = utc_days * 86400;                                        
    var date_info = new Date(utc_value * 1000);
 
    var fractional_day = serial - Math.floor(serial) + 0.0000001;
 
    var total_seconds = Math.floor(86400 * fractional_day);
 
    var seconds = total_seconds % 60;
 
    total_seconds -= seconds;
 
    var hours = Math.floor(total_seconds / (60 * 60));
    var minutes = Math.floor(total_seconds / 60) % 60;
 
    return new Date(date_info.getFullYear(), date_info.getMonth(), date_info.getDate(), hours, minutes, seconds);
 }

//console.log(hiringDataArray)