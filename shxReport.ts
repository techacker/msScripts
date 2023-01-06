//****************************************************************
//****              2023 SHX Hiring Trend SCRIPT            ******
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
    //let [numGlobal, numEMEA, numPL, numCH, numD, numFR, numIN, numIT, numNA, numSA, numMO, numTotal, numPLFilled, numCHFilled, numDFilled, numFRFilled, numINFilled, numITFilled, numNAFilled, numSAFilled, numMOFilled, numLATAMFilled, num2021, num21CHFilled, num21DFilled, num21FRFilled, num21INFilled, num21ITFilled, num21NAFilled, num21CHPosted, num21DPosted, num21FRPosted, num21INPosted, num21ITPosted, num21NAPosted, num21EMEAPosted, num21Filled, num2022, num22PLFilled, num22CHFilled, num22DFilled, num22FRFilled, num22INFilled, num22ITFilled, num22NAFilled, num22SAFilled, num22MOFilled, num22PLPosted, num22CHPosted, num22DPosted, num22FRPosted, num22INPosted, num22ITPosted, num22NAPosted, num22SAPosted, num22MOPosted, num22EMEAPosted, num22Filled, num2023, num23PLFilled, num23CHFilled, num23DFilled, num23FRFilled, num23INFilled, num23ITFilled, num23NAFilled, num23SAFilled, num23MOFilled, num23PLPosted, num23CHPosted, num23DPosted, num23FRPosted, num23INPosted, num23ITPosted, num23NAPosted, num23SAPosted, num23MOPosted, num23EMEAPosted, num23Filled] = countsArray
  
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
          if (entity === item && reqStatus.toLowerCase() !== "canceled" && reqStatus.toLowerCase() !== "cancelled" && entity !== "") {
          // Get data related to SHX org 
          hiringDataArray.push([jobTitle, adpREQ, hmName.trim(), reqStatus, globalHC, headcount, poland, china, germany, france, india, italy, usa, brazil, morocco, candidateName, changeDate(startDate), offerMade, offerPending, priority, changeIntExt(internal), changeYear(targetYear), changeCountries(country), grade, freshGrads, changeEntity(entity)])
          }
      })
    })
  
    // Update data in Report tab
    if (repLR > 1) {
        repss.getRangeByIndexes(1, 0, repLR - 1, repLC).clear()
    }
  
    repss.getRangeByIndexes(1,0,hiringDataArray.length,hiringDataArray[0].length).setValues(hiringDataArray)
  
    return true
    
  }
  
  // Convert Entities to Short form
  function changeEntity(entity:string):string {
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
  
  // Segregate internal and external candidates
  function changeIntExt(internal:string):string {
      // Segregate internal and external candidates
      if (internal == "") {
          internal = 'External'
      } else {
          internal = 'Internal'
      }
      return internal
  }
  
  // Convert countries names from their symbols
  function changeCountries(country:string):string {
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
  
  // Convert non-US Date format to US Date format
  function changeDate(startDate:string | number):string {
      if (typeof(startDate) === "string") {
          let dateSplit = startDate.split('/')
          if (dateSplit.length > 1) {
              // Convert non-US format to US Date format
              let yyyy = ""
              // Convert Year part to 2022, 2023 etc.
              if (dateSplit[2] !== undefined) {
                  if (dateSplit[2].length == 2) {
                      yyyy = '20' + dateSplit[2]
                  } else {
                      yyyy = dateSplit[2]
                  }
              }
              // Check for date being in US format already, if not put month first
              if (Number(dateSplit[0]) > 12) {
                  startDate = parseInt(dateSplit[1]) + '/' + parseInt(dateSplit[0]) + '/' + yyyy
              } else {
                  startDate = parseInt(dateSplit[0]) + '/' + parseInt(dateSplit[1]) + '/' + yyyy
              } 
          }
      }
      else if (typeof(startDate) === "number") {
          // One liner function to convert Excel Serials to US Date format
          startDate = new Date(Date.UTC(0,0,startDate)).toLocaleDateString('en-US')
      }
      else {
          startDate = ""
      }
      return startDate
  }
  
  // Target Year 2021, 2022 and 2023
  function changeYear(targetYear:string):string {
      // Target Year 2021, 2022 and 2023
      if (targetYear == '') {
          targetYear = '2022'
      }
      return targetYear
  }