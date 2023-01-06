function main(workbook: ExcelScript.Workbook) {
    let repSS = workbook.getWorksheet('Report')
    let colCount = repSS.getUsedRange().getColumnCount()
    let rowCount = repSS.getUsedRange().getRowCount()
    let repHeaderRange = repSS.getRangeByIndexes(0, 0, 1, colCount).getValues()[0]  //1st Value of Array
    let repDataRange = repSS.getRangeByIndexes(1, 0, rowCount -1, colCount).getValues()  // Change 5 with 'rowCount -1'
    let countsArray:number[] = []
    let [numGlobal=0, numEMEA=0, numPL=0, numCH=0, numD=0, numFR=0, numIN=0, numIT=0, numNA=0, numSA=0, numMO=0, numTotal=0] = countsArray
    let [numPLFilled=0, numCHFilled=0, numDFilled=0, numFRFilled=0, numINFilled=0, numITFilled=0, numNAFilled=0, numSAFilled=0, numMOFilled=0,numLATAMFilled=0] = countsArray
    //let [num2021=0, num21CHFilled=0, num21DFilled=0, num21FRFilled=0, num21INFilled, num21ITFilled, num21NAFilled, num21CHPosted, num21DPosted, num21FRPosted, num21INPosted, num21ITPosted, num21NAPosted, num21EMEAPosted,num21Filled=0] = countsArray
    let [num2022=0, num22PLFilled=0, num22CHFilled=0, num22DFilled=0, num22FRFilled=0, num22INFilled=0, num22ITFilled=0, num22NAFilled=0, num22SAFilled=0] = countsArray
    let [num22MOFilled=0, num22PLPosted=0, num22CHPosted=0, num22DPosted=0, num22FRPosted=0, num22INPosted=0, num22ITPosted=0, num22NAPosted=0, num22SAPosted=0, num22MOPosted=0, num22EMEAPosted=0] = countsArray
    let [num22Filled=0, num2023=0, num23PLFilled=0, num23CHFilled=0, num23DFilled=0, num23FRFilled=0, num23INFilled=0, num23ITFilled=0, num23NAFilled=0, num23SAFilled=0, num23MOFilled=0] = countsArray
    let [num23PLPosted=0, num23CHPosted=0, num23DPosted=0, num23FRPosted=0, num23INPosted=0, num23ITPosted=0, num23NAPosted=0, num23SAPosted=0, num23MOPosted=0, num23EMEAPosted=0, num23Filled=0] = countsArray

    let HeaderObj = {}
    
    repDataRange.forEach(row => {
        repHeaderRange.forEach((heading:string, colInd:number) => {
            HeaderObj[heading] = row[colInd]
        })
        
        let jobTitle:string = HeaderObj["Job Title"]
        let adpREQ:string = HeaderObj["ADP Req"]
        let hmName:string = HeaderObj["Hiring Manager Name(s)"]
        let reqStatus:string = HeaderObj["Status"]
        let globalHC:number = HeaderObj["Total Global HC Per Job"]
        let headcount:number = HeaderObj["Total HC Per Opening"]
        let poland:boolean = HeaderObj["PL"]
        let china:boolean = HeaderObj["CH"]
        let germany:boolean = HeaderObj["D"]
        let france:boolean = HeaderObj["FR"]
        let india:boolean = HeaderObj["IN"]
        let italy:boolean = HeaderObj["IT"]
        let usa:boolean = HeaderObj["NA"]
        let brazil:boolean = HeaderObj["SA"]
        let morocco:boolean = HeaderObj["MO"]
        let candidateName:string = HeaderObj["Selected Candidate"]
        let startDate:string = HeaderObj["Start Date"]
        let offerMade:string = HeaderObj["Offer Extended"]
        let offerPending:string = HeaderObj["Offer to be made"]
        let priority:string = HeaderObj["Top Priority"]
        let internal:string = HeaderObj["Internal/External"]
        let targetYear:string = HeaderObj["Target Year"]
        let country:string = HeaderObj["Region Hired"]
        let grade:string = HeaderObj["Job Grade"]
        let freshGrads:string = HeaderObj["Fresh Grads"]
        let entity:string = HeaderObj["SHX Entity"]
        let candUpdates:string = HeaderObj["Candidate Updates"]   
        //console.log(`${jobTitle} is ${reqStatus} in ${country} and the headcount is ${headcount}.`)
        if (headcount !== 0) {
            numTotal += headcount
            if (targetYear == "2021" || targetYear == "2022") {
                num2022 += headcount
                if (china == true) { num22CHPosted += headcount }
                if (poland == true) { num22PLPosted += headcount }
                if (germany == true) { num22DPosted += headcount }
                if (france == true) { num22FRPosted += headcount }
                if (india == true) { num22INPosted += headcount }
                if (italy == true) {num22ITPosted += headcount }
                if (usa == true) { num22NAPosted += headcount }
                if (brazil == true) { num22SAPosted += headcount }
                if (morocco == true) { num22MOPosted += headcount }
                if (germany == true || france == true || italy == true) { num22EMEAPosted += headcount }
                if (country == "China") { num22CHFilled += headcount }
                    else if (country == "Poland") { num22PLFilled += headcount }
                    else if (country == "Germany") { num22DFilled += headcount }
                    else if (country == "France") { num22FRFilled += headcount }
                    else if (country == "India") { num22INFilled += headcount } 
                    else if (country == "Italy") { num22ITFilled += headcount } 
                    else if (country == "USA") { num22NAFilled += headcount } 
                    else if (country == "Brazil") { num22SAFilled += headcount } 
                    else if (country == "Morocco") { num22MOFilled += headcount }
            }
            else if (targetYear == "2023") {
                num2023 += headcount
                if (china == true) { num23CHPosted += headcount }
                if (poland == true ) { num23PLPosted += headcount }
                if (germany == true) { num23DPosted += headcount }
                if (france == true) { num23FRPosted += headcount }
                if (india == true) { num23INPosted += headcount }
                if (italy == true) { num23ITPosted += headcount }
                if (usa == true) { num23NAPosted += headcount }
                if (brazil == true) { num23SAPosted += headcount }
                if (morocco == true) { num23MOPosted += headcount }
                if (germany == true || france == true || italy == true) {
                    num23EMEAPosted += headcount
                }
                if (country == "China") { num23CHFilled += headcount } 
                    else if (country == "Poland") { num23PLFilled += headcount } 
                    else if (country == "Germany") { num23DFilled += headcount } 
                    else if (country == "France") { num23FRFilled += headcount } 
                    else if (country == "India") { num23INFilled += headcount } 
                    else if (country == "Italy") { num23ITFilled += headcount } 
                    else if (country == "USA") { num23NAFilled += headcount } 
                    else if (country == "Brazil") { num23SAFilled += headcount } 
                    else if (country == "Morocco") { num23MOFilled += headcount }
            }

            // China Only
            if (china == true && germany == false && france == false && india == false && italy == false && usa == false && brazil == false && poland == false && morocco == false) {
                numCH += headcount
            }
            // Poland Only
            else if (china == false && poland == true && germany == false && france == false && india == false && italy == false && usa == false && brazil == false && morocco == false) {
                numPL += headcount
            }
            // Germany Only
            else if (china == false && germany == true && france == false && india == false && italy == false && usa == false && brazil == false && poland == false && morocco == false) {
                numD += headcount
            }
            // France Only
            else if (china == false && germany == false && france == true && india == false && italy == false && usa == false && brazil == false && poland == false && morocco == false) {
                numFR += headcount
            }
            // India Only
            else if (china == false && germany == false && france == false && india == true && italy == false && usa == false && brazil == false && poland == false && morocco == false) {
                numIN += headcount
            }
            // Italy Only
            else if (china == false && germany == false && france == false && india == false && italy == true && usa == false && brazil == false && poland == false && morocco == false) {
                numIT += headcount
            }
            // US Only
            else if (china == false && germany == false && france == false && india == false && italy == false && usa == true && brazil == false && poland == false && morocco == false) {
                numNA += headcount
            }
            // Morocco Only
            else if (china == false && germany == false && france == false && india == false && italy == false && usa == false && brazil == false && poland == false && morocco == true) {
                numMO += headcount
            }
            // SA/Brazil Only
            else if (china == false && germany == false && france == false && india == false && italy == false && usa == false && brazil == true && poland == false && morocco == false) {
                numSA += headcount
            }
            // EMEA
            else if (france == true && germany == true && italy == true && china == false && india == false && usa == false && brazil == false) {
                numEMEA += headcount
            }
            // Global (ignoring China)
            else if (germany == true && france == true && india == true && italy == true && usa == true && brazil == true) {
                numGlobal += headcount
            }
            else {
                //console.log('Random locations posted.')
            }
        }
    })
    console.log(numEMEA + " " + numGlobal)
}