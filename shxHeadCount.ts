function main(workbook: ExcelScript.Workbook) {
    let repSS = workbook.getWorksheet('Report')
    let colCount = repSS.getUsedRange().getColumnCount()
    let rowCount = repSS.getUsedRange().getRowCount()
    let lastRan:string = new Date().toLocaleDateString()
    let repHeaderRange = repSS.getRangeByIndexes(0, 0, 1, colCount).getValues()[0] 
    let repDataRange = repSS.getRangeByIndexes(1, 0, rowCount-1, colCount).getValues()  // rowCount -1

    // Stats Sheet Info
    let statSS = workbook.getWorksheet('Stats')
    let statLC = statSS.getUsedRange().getColumnCount()
    let statLR = statSS.getUsedRange().getRowCount()
    let statDataRange = statSS.getRangeByIndexes(0,0,statLR-2,statLC)
    
    let countsArray:number[] = []
    let [numGlobal=0, numEMEA=0, numPL=0, numCH=0, numD=0, numFR=0, numIN=0, numIT=0, numNA=0, numSA=0, numMO=0, numTotal=0, numFilled=0, numActive=0, numHCC=0, numLCC=0] = countsArray
    let [numPLFilled=0, numCHFilled=0, numDFilled=0, numFRFilled=0, numINFilled=0, numITFilled=0, numNAFilled=0, numSAFilled=0, numMOFilled=0,numEMEAFilled=0, numLATAMFilled=0] = countsArray
    let [num2022=0, num22PLFilled=0, num22CHFilled=0, num22DFilled=0, num22FRFilled=0, num22INFilled=0, num22ITFilled=0, num22NAFilled=0, num22SAFilled=0] = countsArray
    let [num22MOFilled=0, num22PLPosted=0, num22CHPosted=0, num22DPosted=0, num22FRPosted=0, num22INPosted=0, num22ITPosted=0, num22NAPosted=0, num22SAPosted=0, num22MOPosted=0, num22EMEAPosted=0] = countsArray
    let [num22Filled=0, num2023=0, num23PLFilled=0, num23CHFilled=0, num23DFilled=0, num23FRFilled=0, num23INFilled=0, num23ITFilled=0, num23NAFilled=0, num23SAFilled=0, num23MOFilled=0] = countsArray
    let [num23PLPosted=0, num23CHPosted=0, num23DPosted=0, num23FRPosted=0, num23INPosted=0, num23ITPosted=0, num23NAPosted=0, num23SAPosted=0, num23MOPosted=0, num23EMEAPosted=0, num23EMEAFilled=0, num23Filled=0] = countsArray
    let [num22Active=0, num23Active=0,num22EMEAFilled=0, num22HCC=0, num22LCC=0, num23HCC=0, num23LCC=0, perc22HCC=0, perc22LCC=0, perc23HCC=0, perc23LCC=0, percHCC=0, percLCC=0] = countsArray

    let HeaderObj = {}
    let statData:(string|number)[][] = []
    
    // Calculate 2023 and earlier stats
    repDataRange.forEach(row => {
        repHeaderRange.forEach((heading:string, colInd:number) => {
            HeaderObj[heading] = row[colInd]
        })
        
        let jobTitle:string = HeaderObj["Job Title"]
        let adpREQ:string = HeaderObj["ADP Req"]
        let hmName:string = HeaderObj["Hiring Manager's Names"]
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
        let startDate:(number|string) = HeaderObj["Start Date (US Format)"]
        let offerMade:string = HeaderObj["Offer Made"]
        let offerPending:string = HeaderObj["Offer Pending"]
        let priority:string = HeaderObj["Top Priority"]
        let internal:string = HeaderObj["Internal/External"]
        let targetYear:string = HeaderObj["Target Year"]
        let country:string = HeaderObj["Region Hired"]
        let grade:string = HeaderObj["Job Grade"]
        let freshGrads:string = HeaderObj["Fresh Grads"]
        let entity:string = HeaderObj["SHX Entity"]
        let candUpdates:string = HeaderObj["Candidate Updates"]   
        if (headcount !== 0) {
            //numTotal += headcount
            // Calculate FILLED count per country per year
            if (targetYear == "2021" || targetYear == "2022") {
                num2022 += headcount
                // Calculate FILLED count per country
                let filled22:number[] = calculateFilled(country,headcount)
                num22CHFilled += filled22[0]
                num22DFilled += filled22[1]
                num22FRFilled += filled22[2]
                num22INFilled += filled22[3]
                num22ITFilled += filled22[4]
                num22NAFilled += filled22[5]
                num22SAFilled += filled22[6]
                num22PLFilled += filled22[7]
                num22MOFilled += filled22[8]
                
                // Calculate POSTED count per country
                let posted22:number[] = calculatePosted(china,germany,france,india,italy,usa,brazil,poland,morocco,headcount)
                num22CHPosted += posted22[0]
                num22DPosted += posted22[1]
                num22FRPosted += posted22[2]
                num22INPosted += posted22[3]
                num22ITPosted += posted22[4]
                num22NAPosted += posted22[5]
                num22SAPosted += posted22[6]
                num22PLPosted += posted22[7]
                num22MOPosted += posted22[8]
                num22EMEAPosted += posted22[9]
            }
            else if (targetYear == "2023") {
                num2023 += headcount
                // Calculate FILLED count per country
                let filled23:number[] = calculateFilled(country,headcount)
                num23CHFilled += filled23[0]
                num23DFilled += filled23[1]
                num23FRFilled += filled23[2]
                num23INFilled += filled23[3]
                num23ITFilled += filled23[4]
                num23NAFilled += filled23[5]
                num23SAFilled += filled23[6]
                num23PLFilled += filled23[7]
                num23MOFilled += filled23[8]
                
                // Calculate POSTED count per country
                let posted23:number[] = calculatePosted(china,germany,france,india,italy,usa,brazil,poland,morocco,headcount)
                num23CHPosted += posted23[0]
                num23DPosted += posted23[1]
                num23FRPosted += posted23[2]
                num23INPosted += posted23[3]
                num23ITPosted += posted23[4]
                num23NAPosted += posted23[5]
                num23SAPosted += posted23[6]
                num23PLPosted += posted23[7]
                num23MOPosted += posted23[8]
                num23EMEAPosted += posted23[9]
            }
        }
    })
    
    // Upto 2022 Calculation Steps
    num22EMEAPosted = num22DPosted + num22FRPosted + num22ITPosted
    num22EMEAFilled = num22DFilled + num22FRFilled + num22ITFilled
    num22Filled = num22PLFilled + num22DFilled + num22FRFilled + num22INFilled + num22NAFilled + num22SAFilled + num22MOFilled
    num22Active = num2022 - num22Filled
    num22HCC = num22EMEAFilled + num22NAFilled + num22CHFilled
    num22LCC = num22PLFilled + num22INFilled + num22SAFilled + num22MOFilled
    perc22HCC = num22Filled > 0 ? num22HCC/num22Filled : 0
    perc22LCC = num22Filled > 0 ? num22LCC/num22Filled : 0

    // 2023 Calculation Steps
    num23EMEAPosted = num23DPosted + num23FRPosted + num23ITPosted
    num23EMEAFilled = num23DFilled + num23FRFilled + num23ITFilled
    num23Filled = num23PLFilled + num23DFilled + num23FRFilled + num23INFilled + num23NAFilled + num23SAFilled + num23MOFilled
    num23Active = num2023 - num23Filled
    num23HCC = num23EMEAFilled + num23NAFilled + num23CHFilled
    num23LCC = num23PLFilled + num23INFilled + num23SAFilled + num23MOFilled
    perc23HCC = num23Filled > 0 ? num23HCC/num23Filled : 0
    perc23LCC = num23Filled > 0 ? num23LCC/num23Filled : 0

    // Overall Calculation Steps
    numTotal = num2022 + num2023
    numEMEA = num22EMEAPosted + num23EMEAPosted
    numIN = num22INPosted + num23INPosted
    numNA = num22NAPosted + num23NAPosted
    numSA = num22SAPosted + num23SAPosted
    numPL = num22PLPosted + num23PLPosted
    numMO = num22MOPosted + num23MOPosted
    numEMEAFilled = num22EMEAFilled + num23EMEAFilled
    numINFilled = num22INFilled + num23INFilled
    numNAFilled = num22NAFilled + num23NAFilled
    numSAFilled = num22SAFilled + num23SAFilled
    numPLFilled = num22PLFilled + num23PLFilled
    numMOFilled = num22MOFilled + num23MOFilled
    numFilled = num22Filled + num23Filled
    numActive = num2022 + num2023 - numFilled
    numHCC = numEMEAFilled + numNAFilled + numCHFilled
    numLCC = numPLFilled + numINFilled + numSAFilled + numMOFilled
    percHCC = numFilled > 0 ? numHCC/numFilled : 0
    percLCC = numFilled > 0 ? numLCC/numFilled : 0

    // Create the array to be posted as values
    statData.push([lastRan, num2023, num23CHPosted, num23DPosted, num23FRPosted, num23ITPosted, num23EMEAPosted, num23INPosted,num23NAPosted, num23SAPosted, num23PLPosted, num23MOPosted, num23Active,
        num23Filled, num23CHFilled, num23DFilled, num23FRFilled, num23ITFilled, num23EMEAFilled, num23INFilled, num23NAFilled, num23SAFilled, num23PLFilled, num23MOFilled, perc23HCC, perc23LCC, num2022, num22CHPosted, num22DPosted, num22FRPosted, num22ITPosted, num22EMEAPosted,num22INPosted, num22NAPosted, num22SAPosted, num22PLPosted, num22MOPosted, num22Active, num22Filled,num22CHFilled, num22DFilled, num22FRFilled, num22ITFilled, num22EMEAFilled, num22INFilled, num22NAFilled,num22SAFilled, num22PLFilled, num22MOFilled, perc22HCC, perc22LCC, numTotal, numEMEA, numIN, numNA, numSA, numPL, numMO, numActive, numFilled, numEMEAFilled, numINFilled, numNAFilled,numSAFilled, numPLFilled, numMOFilled, percHCC, percLCC])
    
    // Write array values to the 2023 Stats sheet
    statSS.getRangeByIndexes(statLR,0,statData.length,statData[0].length).setValues(statData)
}

function calculatePosted(china:boolean,germany:boolean,france:boolean,india:boolean,italy:boolean,usa:boolean,brazil:boolean,poland:boolean,morocco:boolean,headcount:number):number[] {
    let countsArray:number[] = []
    let [numGlobal=0, numEMEA=0, numPL=0, numCH=0, numD=0, numFR=0, numIN=0, numIT=0, numNA=0, numSA=0, numMO=0] = countsArray
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
    return [numCH,numD,numFR,numIN,numIT,numNA,numSA,numPL,numMO,numEMEA,numGlobal]
}

function calculateFilled(country:string, headcount:number):number[] {
    let countsArray:number[] = []
    let [numPLFilled=0, numCHFilled=0, numDFilled=0, numFRFilled=0, numINFilled=0, numITFilled=0, numNAFilled=0, numSAFilled=0, numMOFilled=0, numFilled=0] = countsArray
    if (country == "China") { numCHFilled += headcount }
    else if (country == "Poland") { numPLFilled += headcount }
    else if (country == "Germany") { numDFilled += headcount }
    else if (country == "France") { numFRFilled += headcount }
    else if (country == "India") { numINFilled += headcount } 
    else if (country == "Italy") { numITFilled += headcount } 
    else if (country == "USA") { numNAFilled += headcount } 
    else if (country == "Brazil") { numSAFilled += headcount } 
    else if (country == "Morocco") { numMOFilled += headcount }
    numFilled = numPLFilled + numDFilled + numFRFilled + numINFilled + numITFilled + numSAFilled + numMOFilled
    return [numCHFilled,numDFilled,numFRFilled,numINFilled,numITFilled,numNAFilled,numSAFilled,numPLFilled,numMOFilled, numFilled]
}
