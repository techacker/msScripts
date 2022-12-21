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

function calculatePerYearHC(headcount) {
    // Get Total HC posted and filled in which country
    if (headcount !== "") {
        numTotal += parseInt(headcount)
        if (targetYear == "2021") {
            num2021 += parseInt(headcount)
            if (china == "TRUE") {
                num21CHPosted += parseInt(headcount)
            }
            if (germany == "TRUE") {
                num21DPosted += parseInt(headcount)
            }
            if (france == "TRUE") {
                num21FRPosted += parseInt(headcount)
            }
            if (india == "TRUE") {
                num21INPosted += parseInt(headcount)
            }
            if (italy == "TRUE") {
                num21ITPosted += parseInt(headcount)
            }
            if (usa == "TRUE") {
                num21NAPosted += parseInt(headcount)
            }
            if (germany == "TRUE" || france == "TRUE" || italy == "TRUE") {
                num21EMEAPosted += parseInt(headcount)
            }
            if (country == "China") {
                num21CHFilled += parseInt(headcount)
            } else if (country == "Germany") {
                num21DFilled += parseInt(headcount)
            } else if (country == "France") {
                num21FRFilled += parseInt(headcount)
            } else if (country == "India") {
                num21INFilled += parseInt(headcount)
            } else if (country == "Italy") {
                num21ITFilled += parseInt(headcount)
            } else if (country == "USA") {
                num21NAFilled += parseInt(headcount)
            }
        } else if (targetYear == "2022") {
            num2022 += parseInt(headcount)
            if (china == "TRUE") {
                num22CHPosted += parseInt(headcount)
            }
            if (poland == "TRUE") {
                num22PLPosted += parseInt(headcount)
            }
            if (germany == "TRUE") {
                num22DPosted += parseInt(headcount)
            }
            if (france == "TRUE") {
                num22FRPosted += parseInt(headcount)
            }
            if (india == "TRUE") {
                num22INPosted += parseInt(headcount)
            }
            if (italy == "TRUE") {
                num22ITPosted += parseInt(headcount)
            }
            if (usa == "TRUE") {
                num22NAPosted += parseInt(headcount)
            }
            if (brazil == "TRUE") {
                num22SAPosted += parseInt(headcount)
            }
            if (morocco == "TRUE") {
                num22MOPosted += parseInt(headcount)
            }
            if (germany == "TRUE" || france == "TRUE" || italy == "TRUE") {
                num22EMEAPosted += parseInt(headcount)
            }
            if (country == "China") {
                num22CHFilled += parseInt(headcount)
            } else if (country == "Poland") {
                num22PLFilled += parseInt(headcount)
            } else if (country == "Germany") {
                num22DFilled += parseInt(headcount)
            } else if (country == "France") {
                num22FRFilled += parseInt(headcount)
            } else if (country == "India") {
                num22INFilled += parseInt(headcount)
            } else if (country == "Italy") {
                num22ITFilled += parseInt(headcount)
            } else if (country == "USA") {
                num22NAFilled += parseInt(headcount)
            } else if (country == "Brazil") {
                num22SAFilled += parseInt(headcount)
            } else if (country == "Morocco") {
                num22MOFilled += parseInt(headcount)
            }
        } else if (targetYear == "2023") {
            num2023 += parseInt(headcount)
            if (china == "TRUE") {
                num23CHPosted += parseInt(headcount)
            }
            if (poland == "TRUE") {
                num23PLPosted += parseInt(headcount)
            }
            if (germany == "TRUE") {
                num23DPosted += parseInt(headcount)
            }
            if (france == "TRUE") {
                num23FRPosted += parseInt(headcount)
            }
            if (india == "TRUE") {
                num23INPosted += parseInt(headcount)
            }
            if (italy == "TRUE") {
                num23ITPosted += parseInt(headcount)
            }
            if (usa == "TRUE") {
                num23NAPosted += parseInt(headcount)
            }
            if (brazil == "TRUE") {
                num23SAPosted += parseInt(headcount)
            }
            if (morocco == "TRUE") {
                num23MOPosted += parseInt(headcount)
            }
            if (germany == "TRUE" || france == "TRUE" || italy == "TRUE") {
                num23EMEAPosted += parseInt(headcount)
            }
            if (country == "China") {
                num23CHFilled += parseInt(headcount)
            } else if (country == "Poland") {
                num23PLFilled += parseInt(headcount)
            } else if (country == "Germany") {
                num23DFilled += parseInt(headcount)
            } else if (country == "France") {
                num23FRFilled += parseInt(headcount)
            } else if (country == "India") {
                num23INFilled += parseInt(headcount)
            } else if (country == "Italy") {
                num23ITFilled += parseInt(headcount)
            } else if (country == "USA") {
                num23NAFilled += parseInt(headcount)
            } else if (country == "Brazil") {
                num23SAFilled += parseInt(headcount)
            } else if (country == "Morocco") {
                num23MOFilled += parseInt(headcount)
            }
        }
    }
}

function calculateRegionalData(headcount) {
    // Get regional data
    if (headcount !== "") {
        // China Only
        if (china == "TRUE" && germany == "FALSE" && france == "FALSE" && india == "FALSE" && italy == "FALSE" && usa == "FALSE" && brazil == "FALSE" && poland == "FALSE" && morocco == "FALSE") {
            numCH += parseInt(headcount)
        } 
        // Poland Only
        else if (china == "FALSE" && poland == "TRUE" && germany == "FALSE" && france == "FALSE" && india == "FALSE" && italy == "FALSE" && usa == "FALSE" && brazil == "FALSE" && morocco == "FALSE") {
            numPL += parseInt(headcount)
        }
        // Germany Only
        else if (china == "FALSE" && germany == "TRUE" && france == "FALSE" && india == "FALSE" && italy == "FALSE" && usa == "FALSE" && brazil == "FALSE" && poland == "FALSE" && morocco == "FALSE") {
            numD += parseInt(headcount)
        }
        // France Only
        else if (china == "FALSE" && germany == "FALSE" && france == "TRUE" && india == "FALSE" && italy == "FALSE" && usa == "FALSE" && brazil == "FALSE" && poland == "FALSE" && morocco == "FALSE") {
            numFR += parseInt(headcount)
        }
        // India Only
        else if (china == "FALSE" && germany == "FALSE" && france == "FALSE" && india == "TRUE" && italy == "FALSE" && usa == "FALSE" && brazil == "FALSE" && poland == "FALSE" && morocco == "FALSE") {
            numIN += parseInt(headcount)
        }
        // Italy Only
        else if (china == "FALSE" && germany == "FALSE" && france == "FALSE" && india == "FALSE" && italy == "TRUE" && usa == "FALSE" && brazil == "FALSE" && poland == "FALSE" && morocco == "FALSE") {
            numIT += parseInt(headcount)
        }
        // US Only
        else if (china == "FALSE" && germany == "FALSE" && france == "FALSE" && india == "FALSE" && italy == "FALSE" && usa == "TRUE" && brazil == "FALSE" && poland == "FALSE" && morocco == "FALSE") {
            numNA += parseInt(headcount)
        }
        // Morocco Only
        else if (china == "FALSE" && germany == "FALSE" && france == "FALSE" && india == "FALSE" && italy == "FALSE" && usa == "FALSE" && brazil == "FALSE" && poland == "FALSE" && morocco == "TRUE") {
            numMO += parseInt(headcount)
        }
        // SA/Brazil Only
        else if (china == "FALSE" && germany == "FALSE" && france == "FALSE" && india == "FALSE" && italy == "FALSE" && usa == "FALSE" && brazil == "TRUE" && poland == "FALSE" && morocco == "FALSE") {
            numSA += parseInt(headcount)
        }
        // EMEA
        else if (france == "TRUE" && germany == "TRUE" && italy == "TRUE" && china == "FALSE" && india == "FALSE" && usa == "FALSE" && brazil == "FALSE") {
            numEMEA += parseInt(headcount)
        }
        // Global (irrespective of Brazil and China)
        else if (germany == "TRUE" && france == "TRUE" && india == "TRUE" && italy == "TRUE" && usa == "TRUE") {// && brazil== "TRUE") {
            numGlobal += parseInt(headcount)
        }
    }
}

console.log(hiringDataArray)

/*
function main(workbook: ExcelScript.Workbook): EventData[] {
    let table = workbook.getWorksheet("Sheet1").getTables()[0];
    let range = table.getRangeBetweenHeaderAndTotal();
    let rows = range.getValues();
    let records: EventData[] = [];

    for (let row of rows) {
        let [JobTitle, YearsOfExperience, EmpID] = row;
        records.push({
            JobTitle: JobTitle as string,
    YearsOfExperience: YearsOfExperience as number,
    EmpID: EmpID as number
        })
    }
    console.log(JSON.stringify(records));
    return records;
}
interface EventData {
    JobTitle: string
    YearsOfExperience: number
    EmpID: number
}
*/