//****************************************************************
//****            2023 SHX Hiring Trend SCRIPT              ******
//****               AUTHOR: ANURAG BANSAL                  ******
//****                Revision: 1.0.0                       ******
//****                Date: 11/14/2022                      ******
//****************************************************************
//****                Revision History                      ******
//****************************************************************

function main(workbook: ExcelScript.Workbook) {
    
  // Define the Data Reference Sheets
  const nnss = workbook.getWorksheet("New Needs")
  const repss = workbook.getWorksheet("Report")
  const dateRan = new Date().toLocaleDateString()
  const shxdept = ["1-SW Engineering SHX/SWE", "2-SW Projects  SHX/ADPM", "16-EE & HW Experience SHX/EEHW", "6-ADAS SHX/ADX", "10-SW Artificial Intelligence SHX/SAI", "11-User Experience SHX/UEXP", "13- Cockpit Connected Services CCS"]

  // Define the data ranges
  const nnRange = nnss.getUsedRange()
  const nnLR = nnRange.getRowCount()
  const nnLC = nnRange.getColumnCount()
  const nnHeaders = nnss.getRangeByIndexes(0,0,1,nnLC).getValues()[0]
  const nnDataArray = nnss.getRangeByIndexes(1,0,nnLR-1,nnLC).getValues()

  // Define the Report Tab Ranges
  const repLR = repss.getUsedRange().getRowCount()
  const repLC = repss.getUsedRange().getColumnCount()

  let HeaderObj = {}
  let hiringDataArray: (string | number | boolean)[][] = []
  let countsArray:number[] = []
  let [numGlobal, numEMEA, numPL, numCH, numD, numFR, numIN, numIT, numNA, numSA, numMO, numTotal, numPLFilled, numCHFilled, numDFilled, numFRFilled, numINFilled, numITFilled, numNAFilled, numSAFilled, numMOFilled, numLATAMFilled, num2021, num21CHFilled, num21DFilled, num21FRFilled, num21INFilled, num21ITFilled, num21NAFilled, num21CHPosted, num21DPosted, num21FRPosted, num21INPosted, num21ITPosted, num21NAPosted, num21EMEAPosted, num21Filled, num2022, num22PLFilled, num22CHFilled, num22DFilled, num22FRFilled, num22INFilled, num22ITFilled, num22NAFilled, num22SAFilled, num22MOFilled, num22PLPosted, num22CHPosted, num22DPosted, num22FRPosted, num22INPosted, num22ITPosted, num22NAPosted, num22SAPosted, num22MOPosted, num22EMEAPosted, num22Filled, num2023, num23PLFilled, num23CHFilled, num23DFilled, num23FRFilled, num23INFilled, num23ITFilled, num23NAFilled, num23SAFilled, num23MOFilled, num23PLPosted, num23CHPosted, num23DPosted, num23FRPosted, num23INPosted, num23ITPosted, num23NAPosted, num23SAPosted, num23MOPosted, num23EMEAPosted, num23Filled] = countsArray

  nnHeaders.forEach((item:string, ind:number) => {
    HeaderObj[item] = ind
  })

  nnDataArray.forEach(rec => {
    nnHeaders.forEach((heading: string, ind:number) => {
      HeaderObj[heading] = rec[ind]
    })
    let entity:string = HeaderObj["SWX Entity"]
    let targetYear:string = HeaderObj["Target Year New"]
    let hmName:string = HeaderObj["Hiring Manager Name(s)"]
    let jobTitle:string = HeaderObj["Job Title"]
    let reqStatus:string = HeaderObj["Req Status"]
    let adpREQ:string = HeaderObj["ADP Req"]
    let candidateName:string = HeaderObj["Selected Candidate"]
    let startDate:string = HeaderObj["Start Date"]
    let offerMade:string = HeaderObj["Offer Extended"]
    let offerPending:string = HeaderObj["Offer to be made"]
    let priority:string = HeaderObj["Top Priority"]
    let internal:string = HeaderObj["Internal"]
    let poland:string = HeaderObj["PL"]
    let china:string = HeaderObj["CH"]
    let germany:string = HeaderObj["D"]
    let france:string = HeaderObj["FR"]
    let india:string = HeaderObj["IN"]
    let italy:string = HeaderObj["IT"]
    let usa:string = HeaderObj["NA"]
    let brazil:string = HeaderObj["SA"]
    let morocco:string = HeaderObj["MO"]
    let globalHC:string = HeaderObj["Total Global HC Per Job"]
    let headcount:string = HeaderObj["Total Global HC Per Opening"]
    let country:string = HeaderObj["Country"]
    let grade:string = HeaderObj["Job Grade - ex FCA"]
    let freshGrads:string = HeaderObj["Fresh Grads"]

    shxdept.forEach(item => {
      if (entity === item && reqStatus !== "Canceled" && entity !== "") {
          // Convert Entities to Short form 
          switch (entity) {
              case "1-SW Engineering SHX/SWE":
              entity = "SWE"
              break;
              case "2-SW Projects  SHX/ADPM":
              entity = "APDM"
              break;
              case "6-ADAS SHX/ADX":
              entity = "ADX"
              break;
              case "13- Cockpit Connected Services CCS":
              entity = "CCS"
              break;
              case "16-EE & HW Experience SHX/EEHW":
              entity = "EEHW"
              break;
              case "10-SW Artificial Intelligence SHX/SAI":
              entity = "SAI"
              break;
              case "11-User Experience SHX/UEXP":
              entity = "UEXP"
              break;
          }
        
        // Segregate internal and external candidates
        if (internal == "") {
            internal = "External"
            } else {
                internal = "Internal"
        }

        // convert countries names from their symbols
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

        /*
        // Convert Start Date to US Format
        if (startDate !== "") {
            console.log(startDate)
            let sdsplit:string[] = startDate.split('/')
            console.log(sdsplit)
            let yyyy:string = ""
            // Convert Year part to 2022, 2023 etc.
            if (sdsplit[2] !== undefined) {
                if (sdsplit[2].length == 2) {
                    yyyy = '20' + sdsplit[2]
                } else {
                    yyyy = sdsplit[2]
                }
            }
            // Check for date being in US format already, if not put month first
            if (Number(sdsplit[1]) > 12) {
                startDate = sdsplit[0] + '/' + sdsplit[1] + '/' + yyyy
            } else {
                startDate = sdsplit[1] + '/' + sdsplit[0] + '/' + yyyy
            } 
        }
        */
        
        // Target Year 2021 & 2022
        if (targetYear === "") {
            targetYear = "2022"
        }

        // Get data related to SWE org only
        hiringDataArray.push([jobTitle, adpREQ, hmName.trim(), reqStatus, globalHC, headcount, poland, china, germany, france, india, italy, usa, brazil, morocco, candidateName, startDate, offerMade, offerPending, priority, internal, targetYear, country, grade, freshGrads, entity])

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
    })
  })

  // Update data in Report tab
  if (repLR > 1) {
      repss.getRangeByIndexes(1, 0, repLR - 1, repLC).clear()
  }
  
  repss.getRangeByIndexes(1,0,hiringDataArray.length,hiringDataArray[0].length).setValues(hiringDataArray)
  
}
