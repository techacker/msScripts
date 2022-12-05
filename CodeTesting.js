// Objective Test

let HeaderObj = {}
let shxDept = ['1-SW Engineering SHX/SWE','6-ADAS SHX/ADX','2-SW Projects  SHX/ADPM','CCS']
let nnHeaders = ["Year","Name","Joining","Dept","Country"]
let nnDataArray = [['2021','Anurag','01/18/2022','1-SW Engineering SHX/SWE','NA'],['','Shipra','09/18/2022','2-SW Projects  SHX/ADPM','BR']]
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
    let country = HeaderObj['Country']

    //console.log(changeEntity(entity))
    hiringDataArray.push([changeYear(targetYear),candName,startDate,changeEntity(entity),changeCountries(country)])
})

function changeEntity(entity) {
    // Convert Entities to Short form
    if (entity == '1-SW Engineering SHX/SWE') {
        entity = 'SWE'
    } else if (entity == '2-SW Projects  SHX/ADPM') {
        entity = 'APDM'
    } else if (entity == '6-ADAS SHX/ADX') {
        entity = 'ADX'
    } else if (entity == '13- Cockpit Connected Services CCS') {
        entity = 'CCS'
    } else if (entity == '16-EE & HW Experience SHX/EEHW') {
        entity = 'EEHW'
    } else if (entity == '10-SW Artificial Intelligence SHX/SAI') {
        entity = 'SAI'
    } else if (entity == '11-User Experience SHX/UEXP') {
        entity = 'UEXP'
    } else {
        entity = 'Error'
    }
    return entity
}

function changeYear(targetYear) {
    // Target Year 2021, 2022 and 2023
    if (targetYear == '') {
        targetYear = '2023'
    }
    return targetYear
}

function changeCountries(country) {
    // Convert countries names from their symbols
    if (country.trim() == "CH") {
        country = "China"
        } else if (country.trim() == "PL") {
            country = "Poland"
        } else if (country.trim() == "BR") {
            country = "Brazil"
        } else if (country.trim() == "D") {
            country = "Germany"
        } else if (country.trim() == "FR") {
            country = "France"
        } else if (country.trim() == "IT") {
            country = "Italy"
        } else if (country.trim() == "IND") {
            country = "India"
        } else if (country.trim() == "NA") {
            country = "USA"
        } else if (country.trim() == "M") {
            country = "Morocco"
    }
    return country
}

function changeIntExt(internal) {
    // Segregate internal and external candidates
    if (internal == "") {
        internal = 'External'
    } else {
        internal = 'Internal'
    }
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

console.log(hiringDataArray)