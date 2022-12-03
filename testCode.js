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

console.log(hiringDataArray)
