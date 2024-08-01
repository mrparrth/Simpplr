function App() {
  this.ss = SpreadsheetApp.getActive()
  this.settings = _getSettings_()

  this.initialize = () => {
    if (this.settings.doInitialReset && this.ss.getId() !== this.settings.templateFile) { //does a reset if doInitialReset is checked and is not the template
      this.ss.getSheetByName('⚙️Meta').getRange('localDatabase').setValue('')
      this.ss.getSheetByName('⚙️Meta').getRange('doInitialReset').setValue(false)
      this.settings.localDatabase = ''
    }
  }
  this.initialize()

  this.reportTemplateName = 'Report Template'
  this.reportRecoTemplate = 'Report Reco Template'
  this.gainsightSheetName = 'INPUT_Gainsight'

  this.showPlanningSheet = () => _showOnlySheet_(this.ss.getSheetByName('Planning'))
  this.showHomeSheet = () => _showOnlySheet_(this.ss.getSheetByName('Home'))
  this.showGainsightSheet = () => _showOnlySheet_(this.ss.getSheetByName(this.gainsightSheetName))


  this.importFromCentralDb = (dbSheetName, json = true) => {
    if (!this.ssCentral) this.ssCentral = SpreadsheetApp.openById(this.settings.centralDatabase)
    let dbSheet = this.ssCentral.getSheetByName(dbSheetName)

    let dbData
    if (json) {
      dbData = _getItemsFromSheet_(dbSheet, row => !!row.key)
    } else {
      dbData = _generateKeyValuePairFromSheet_(dbSheet, 1, 1, 2)
    }

    _toast_(`Data fetched from ${dbSheetName}!`)
    return dbData
  }

  this.resetDatabases = (prompt = true) => {
    if (prompt && !_confirm_('This will reset this sheet and start afresh. Do you want to continue?')) return

    let goals = this.importFromCentralDb('⚙️Goals')
    let objectives = this.importFromCentralDb('⚙️Objectives')
    let solutions = this.importFromCentralDb('⚙️Solutions')
    let recos = this.importFromCentralDb('⚙️Recommendations')
    let urls = this.importFromCentralDb('⚙️Urls')
    let exCorePillars = this.importFromCentralDb('⚙️Ex Core Pillars', false)

    solutions.forEach(solution => {
      let relRecos = recos.filter(reco => reco.linkToSolution == `${solution.exCorePillar} | ${solution.key}`)
      relRecos.forEach(reco => {
        let relUrls = urls.filter(url => url.recommendation == reco.key)
        reco.helpUrls = relUrls.filter(({ type }) => type == '❓Help')
        reco.bestPractices = relUrls.filter(({ type }) => type == '🅱️Best Practices')
      })

      solution.recos = relRecos
    })

    let json = { goals, objectives, solutions, exCorePillars }

    this.saveDbData(json)

    this.ss.getSheetByName('⚙️Meta').getRange(2, 2, 3, 2).clearContent()
    // let localdbFile = this.getLocalDbFile()
    // localdbFile.setContent(JSON.stringify(json))
    // merge gainsight data while resetting?
    // let gainsightData = this.ss.getSheetByName(this.gainsightSheetName).getDataRange().getValues()
    // if (gainsightData.length > 1)
    //   this.mergeGainsightWithDatabase(false)
  }

  this.showGoalsPrompt = (exCorePillar) => {
    let template = HtmlService.createTemplateFromFile('goals')
    let data = this.getDbObject()

    template.data = data
    template.exCorePillar = exCorePillar

    let htmlOutput = template.evaluate()

    _openDialog_(htmlOutput, `Set Goals for ${data.exCorePillars[exCorePillar]} (Select 3-5 KPIs)`, 1000, 520)
  }

  this.showSolutionsPrompt = (exCorePillar) => {
    let data = this.getDbObject()

    let template = HtmlService.createTemplateFromFile('solutions')
    template.data = data
    template.exCorePillar = exCorePillar
    template.platform = this.settings.platform

    let htmlOutput = template.evaluate()

    _openDialog_(htmlOutput, `${data.exCorePillars[exCorePillar]} : Select Business Goals (Next 3-6 Months)`, 1700, 799)
  }

  this.showObjectivesPrompt = (exCorePillar) => {
    let data = this.getDbObject()

    let template = HtmlService.createTemplateFromFile('objectives')
    template.data = data
    template.exCorePillar = exCorePillar

    let htmlOutput = template.evaluate()

    _openDialog_(htmlOutput, `Objectives for ${data.exCorePillars[exCorePillar]}`, 1000, 500)
  }

  this.saveSolutions = (data, exCorePillar) => {
    // let file = DriveApp.createFile('Test.txt', JSON.stringify(data))
    // console.log(file.getUrl())

    let newSolutions = []
    let newRecos = []
    if (!this.ssCentral) this.ssCentral = SpreadsheetApp.openById(this.settings.centralDatabase)

    //extract the manual recos and solutions
    let relatedSolutions = data.solutions.filter(row => row.exCorePillar == exCorePillar)

    for (let solution of relatedSolutions) {
      if (solution.manual) {
        delete solution.manual
        newSolutions.push([solution.title, solution.description, solution.exCorePillar, solution.importance, this.ss.getUrl(), new Date()])
      }

      for (let reco of solution.recos) {
        if (reco.manual) {
          delete reco.manual
          newRecos.push([reco.title, reco.description, solution.key, exCorePillar, reco.csmOrProductData, this.ss.getUrl(), new Date()])
        }
      }
    }

    //add selected values to meta
    let splitExPillars = this.settings.exCorePillars.split('|')
    let exCorePillarName = data.exCorePillars[exCorePillar]
    let selectedPillarIndex = splitExPillars.indexOf(exCorePillarName)

    if (selectedPillarIndex !== -1) {
      let solutionsSelected = relatedSolutions.filter(solution => solution.select).length
      let recosSelected = 0
      relatedSolutions.filter(solution => solution.select).forEach(solution => solution.recos.forEach(reco => reco.addToSuccessPlan ? recosSelected++ : null))
      // relatedSolutions.forEach(solution => solution.recos.forEach(reco => reco.addToSuccessPlan ? recosSelected++ : null))
      this.ss.getSheetByName('⚙️Meta').getRange(selectedPillarIndex + 2, 2, 1, 2).setValues([[solutionsSelected, recosSelected]])
    } else {
      console.error(`Ex Core Pillar ${exCorePillarName} doesn't match with central`)
    }

    //save db
    this.saveDbData(data)

    let manualSolutionSheet = this.ssCentral.getSheetByName('_Manual_Solutions') || this.ssCentral.insertSheet('_Manual_Solutions')
    let manualRecoSheet = this.ssCentral.getSheetByName('_Manual_Recommendations') || this.ssCentral.insertSheet('_Manual_Recommendations')

    if (newSolutions.length > 0) manualSolutionSheet.getRange(manualSolutionSheet.getLastRow() + 1, 1, newSolutions.length, newSolutions[0].length).setValues(newSolutions)
    if (newRecos.length > 0) manualRecoSheet.getRange(manualRecoSheet.getLastRow() + 1, 1, newRecos.length, newRecos[0].length).setValues(newRecos)

    _toast_(`Data saved!`)
  }

  this.protectWorksheet = (sheet, rangesToUnProtect) => {
    let sheetprotections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);

    let protection = sheetprotections.length == 0 ? sheet.protect().setWarningOnly(true) : sheetprotections[0]

    // SpreadsheetApp.getActive().getSheetByName().getProtections()[0].
    let rangesUnprotected = protection.getUnprotectedRanges()
    protection.setUnprotectedRanges([...rangesUnprotected, ...rangesToUnProtect]);
  }

  this.generateReport = () => {
    this.showPlanningSheet()

    _tryAction_(() => this.ss.deleteSheet(this.ss.getSheetByName(this.settings.reportName)))

    let reportTemplate = this.ss.getSheetByName(this.reportTemplateName)
    let recoTemplateSheet = this.ss.getSheetByName(this.reportRecoTemplate)

    let recoTemplateRng = recoTemplateSheet.getRange(`A2:M6`)
    let recoTemplateRngRowsCnt = recoTemplateRng.getNumRows()
    let recoTemplateRngRowHeights = _getRowHeights_(recoTemplateRng)

    let newReport = reportTemplate.copyTo(this.ss)
    newReport.setName(this.settings.reportName).setTabColor(this.settings.reportTabColor)
    _showOnlySheet_(newReport)

    let { goals, objectives, solutions, exCorePillars } = this.getDbObject()
    let allRecos = []
    for (let { recos, select } of solutions) {
      if (!select) continue
      allRecos = [...allRecos, ...recos.filter(row => row.wave !== 'Remove From Plan' && row.addToSuccessPlan)]
    }

    allRecos.sort((reco1, reco2) => reco1.linkToSolution.localeCompare(reco2.linkToSolution))

    //loop over all pillars
    let reactionTexts = this.settings.reactions.split('|')

    let rows = [23, 15, 7]
    let colors = ['#ffe599', '#a4c2f4', '#fce5cd']
    let lastSolution
    let currentSolution, initialRecoRow, coreColor

    (['3-maximize-employee-productivity', '2-engaging-culture-experience', '1-effective-communication']).forEach((exCorePillar, index) => {
      //in case exCorePillar changes when the solution changes, we want to add the header as well
      if (!!lastSolution) {
        //add the solution header to sheet
        newReport.insertRowBefore(initialRecoRow)
        let foundSolution = solutions.find(solution => solution.key == lastSolution)
        newReport.getRange(initialRecoRow, 1, 1, 13).mergeAcross().setValue(foundSolution.title).setBackground(coreColor)
        currentSolution = null
        lastSolution = null
      }

      coreColor = colors[index]
      let relatedGoals = goals.filter(goal => goal.exCorePillar == exCorePillar && goal.select)
      let relatedRecos = allRecos.filter(reco => reco.linkToSolution.startsWith(`${exCorePillar} | `))
      let relatedObjs = objectives.filter(objective => objective.exCorePillar == exCorePillar)
      let startRow = rows[index]

      //add the objectives
      let objectiveOutput = relatedObjs.map(({ title }) => [title])

      if (objectiveOutput.length > 1) newReport.insertRowsAfter(startRow, objectiveOutput.length - 1)
      newReport.getRange(startRow - 1, 4, objectiveOutput.length + 1, 7).setBorder(null, null, null, null, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID)
      newReport.getRange(startRow - 1, 4, 1, 7).setBackground(coreColor).setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
      newReport.getRange(startRow - 1, 4, objectiveOutput.length + 1, 7).setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)

      if (objectiveOutput.length > 0) {
        newReport.getRange(startRow, 4, objectiveOutput.length, 1).setValues(objectiveOutput)
        newReport.getRange(startRow, 4, objectiveOutput.length, 7).mergeAcross()
      }
      startRow = startRow + Math.max(objectiveOutput.length + 2, 3) //+ 2 - objectiveOutput.length

      //add the goals
      let goalOutput = relatedGoals.map(({ title, last30DaysFormatted, '+/-': diff, targetFormatted }) => [title, null, last30DaysFormatted, diff, targetFormatted])

      newReport.insertRowsBefore(startRow, goalOutput.length + 1)
      newReport.getRange(startRow - 1, 4, goalOutput.length + 1, 2).mergeAcross()
      newReport.getRange(startRow - 1, 8, goalOutput.length + 1, 3).mergeAcross()
      newReport.getRange(startRow - 1, 4, goalOutput.length + 1, 7).setBorder(null, null, null, null, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID)
      newReport.getRange(startRow - 1, 4, 1, 7).setBackground(coreColor).setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
      newReport.getRange(startRow - 1, 4, goalOutput.length + 1, 7).setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
      if (goalOutput.length > 0) {
        newReport.getRange(startRow, 4, goalOutput.length, goalOutput[0].length).setValues(goalOutput)
      }

      //add reco to sheet
      let recoStartRow = startRow + goalOutput.length + 2
      initialRecoRow = recoStartRow
      for (let { key, title, description, linkToSolution, note, helpUrls, bestPractices, reaction, owner, due, wave } of relatedRecos) {
        currentSolution = linkToSolution.split(' | ')[1]
        if (currentSolution !== lastSolution && !!lastSolution) {
          //add the solution header to sheet
          newReport.insertRowBefore(initialRecoRow)
          let foundSolution = solutions.find(solution => solution.key == lastSolution)
          newReport.getRange(initialRecoRow, 1, 1, 13).mergeAcross().setValue(foundSolution.title).setBackground(coreColor)
          newReport.getRange(initialRecoRow, 1).setFontWeight('bold').setFontSize(12)
          recoStartRow++
          initialRecoRow = recoStartRow
        }

        newReport.insertRowsBefore(recoStartRow, recoTemplateRngRowsCnt + 1)

        let targetRange = newReport.getRange(recoStartRow, 1, recoTemplateRngRowsCnt, recoTemplateRng.getNumColumns())
        _setRowHeights_(targetRange, recoTemplateRngRowHeights)

        recoTemplateRng.copyTo(targetRange)

        let reactinStr = reaction === '' ? reactionTexts[0] : reactionTexts[reaction + 1]

        let bold = SpreadsheetApp.newTextStyle().setBold(true).build()
        let newRichValueTitle = SpreadsheetApp.newRichTextValue()
          .setText(`${title} - ${description}`)
          .setTextStyle(0, title.length, bold)
          .build()
        let noteRichValue = SpreadsheetApp.newRichTextValue().setText(note).build()

        // newReport.getRange(recoStartRow + 1, 3, 2, 1).setValues([[`${title} - ${description}`], [note]])
        newReport.getRange(recoStartRow + 1, 3, 2, 1).setRichTextValues([[newRichValueTitle], [noteRichValue]])

        let filteredHelps = helpUrls.filter(url => url.platform == this.settings.platform).map(url => url.url)
        let filteredBPs = bestPractices.filter(url => url.platform == this.settings.platform).map(url => url.url)

        newReport.getRange(recoStartRow + 1, 7, 3, 1).setValues([[owner], [_formatDate_(due)], [wave]])
        newReport.getRange(recoStartRow + 1, 11, 3, 1).setValues([[reactinStr], [filteredHelps.join('\n')], [filteredBPs.join('\n')]])

        newReport.getRange(recoStartRow, 14, recoTemplateRngRowsCnt, 2).setValues(new Array(recoTemplateRngRowsCnt).fill([currentSolution, key]))

        this.protectWorksheet(newReport, [
          newReport.getRange(recoStartRow + 1, 7, 3, 1),  // owner, due, wave
          newReport.getRange(recoStartRow + 1, 11),        // reaction
          newReport.getRange(recoStartRow + 2, 3)          // note
        ]);

        recoStartRow = recoStartRow + recoTemplateRngRowsCnt + 1
        lastSolution = currentSolution
      }
    })

    newReport.getRange('K2').setValue(new Date())
    // this.protectWorksheet(newReport, rangesToUnProtect)
  }

  this.saveReportChanges = () => {
    if (!_confirm_('Please confirm to save the changes in this report to database')) return

    let shReport = this.ss.getSheetByName(this.settings.reportName)
    let dbData = this.getDbObject()
    let solutions = dbData.solutions

    //get data to save in db
    let reportdata = shReport.getRange(`A:O`).getDisplayValues().filter(row => !!row[13] && !!row[14])

    let reactionTexts = this.settings.reactions.split('|')

    reportdata.forEach(row => {
      let relSolution = solutions.find(solution => solution.key == row[13])
      let relReco = relSolution.recos.find(reco => reco.key == row[14])

      if (row[1] == 'Note') relReco.note = row[2]
      if (row[5] == 'Owner') relReco.owner = row[6]
      if (row[5] == 'Due') relReco.due = row[6]
      if (row[5] == 'Wave') relReco.wave = row[6]
      if (row[9] == 'Reaction') relReco.reaction = reactionTexts.indexOf(row[10]) == 0 ? '' : reactionTexts.indexOf(row[10]) - 1
    })

    this.saveDbData(dbData)

    _toast_(`Changes in report is successfully saved to recommendation database`)
  }

  this.mergeGainsightWithDatabase = (prompt = true) => {
    if (prompt && !_confirm_('This will update database with data from Gainsight\nDo you want to move forward?')) return

    let gainsightData = _getItemsFromSheet_(this.ss.getSheetByName(this.gainsightSheetName), row => !!row.goalDescription)
    let data = this.getDbObject()

    //clear properties, so we only have new data
    let propsToClear = ['target', 'benchMark', 'last30Days', '+/-', 'targetFormatted', 'benchMarkFormatted', 'last30DaysFormatted']
    data.goals.forEach(goal => {
      for (let prop of propsToClear) {
        delete goal[prop]
      }
    })

    //assign new data to old data
    for (let gainsightRow of gainsightData) {
      let foundGoal = data.goals.find(goal => goal.title == gainsightRow.goalDescription)
      delete gainsightRow['_rowIndex']  //we don't want to merge these 2 properties
      delete gainsightRow['goalDescription']

      if (foundGoal) foundGoal = Object.assign(foundGoal, gainsightRow)
    }

    this.saveDbData(data)

    this.showHomeSheet()
  }

  this.getDbObject = () => {
    let localdbFile = this.getLocalDbFile(false)
    if (!localdbFile) {
      this.resetDatabases(false)
      localdbFile = this.getLocalDbFile()
    }

    let dbData = JSON.parse(localdbFile.getBlob().getDataAsString())
    return dbData
  }

  this.saveDbData = (data) => {
    let localdbFile = this.getLocalDbFile()
    if (!localdbFile) {
      this.resetDatabases(false)
      localdbFile = this.getLocalDbFile()
    }

    localdbFile.setContent(JSON.stringify(data))
  }

  this.getLocalDbFile = (createIfNotPresent = true) => {
    let dbfileid = this.settings.localDatabase
    if (!dbfileid && createIfNotPresent) {
      let prompt = _prompt_('Looks like you are trying to create a plan for a new client. \nPlease enter the name of the client')
      if (prompt.getResponseText() == '' && prompt.getSelectedButton() == SpreadsheetApp.getUi().Button.CANCEL)
        throw 'Please enter a client name and try again!'

      let clientName = prompt.getResponseText()
      let rootFolder = _getFolderById_(this.settings.rootDriveFolder)
      let configFolder = _getFolderByName_('_User Database_', rootFolder, true)
      let date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')

      let file = configFolder.createFile(`${clientName}_${date}.txt`, '{}')

      this.ss.getSheetByName('⚙️Meta').getRange('localDatabase').setValue(file.getId())
      this.settings.localDatabase = file.getId()

      return file
    } else if (dbfileid) {
      return DriveApp.getFileById(dbfileid)
    } else {
      return
    }
  }

  this.exportSheet = (exportFormat = 'pdf') => {
    let activeSheet = this.ss.getActiveSheet()
    if (activeSheet.getName() !== this.settings.reportName) {
      _alert_(`⛔️Please open the report sheet before trying to export!`)
      return
    }

    let sheetParam = '&gid=' + this.ss.getSheetByName(this.settings.reportName).getSheetId()

    let exportUrl = this.ss.getUrl().replace(/\/edit.*$/, '')
      + `/export?exportFormat=${exportFormat}&format=${exportFormat}`
      + '&size=LETTER'
      + `&portrait=true`
      + '&top_margin=0.75'
      + '&bottom_margin=0.75'
      + '&left_margin=0.7'
      + '&right_margin=0.7'
      + '&sheetnames=false&printtitle=false'
      + '&pagenum=CENTER' // change it to CENTER to print page numbers
      + '&gridlines=false'
      + sheetParam

    _openLink_(exportUrl)
  }

  this.checkMigration = () => {
    if (this.ss.getId() == this.templateFile) {
      _alert_(`The migration must start from the client file.`)
      return
    }

    let templateSs = SpreadsheetApp.openById(this.settings.templateFile)
    let currVersion = this.settings.version
    let templateVersion = templateSs.getSheetByName('⚙️Settings').getRange('templateVersion').getValue()

    if (currVersion == templateVersion) {
      _alert_(`⛔️Current version matches with the template version!`)
      return
    }

    if (!_confirm_(`This will create a new file with all your data intact. \nThis file will be trashed. \nDo you want to move ahead?`)) return

    try {
      let driveTemplate = DriveApp.getFileById(this.settings.templateFile)
      let newFile = driveTemplate.makeCopy(this.replaceVersion(this.ss.getName(), templateVersion))
      console.log(newFile.getUrl())
      let folder = _getCurrentFolder_(this.ss.getId())
      newFile.moveTo(folder)

      let newSs = SpreadsheetApp.openById(newFile.getId())
      newSs.getSheetByName('⚙️Settings').getRange('localDatabase').setValue(this.settings.localDatabase)
      newSs.getSheetByName('⚙️Settings').getRange('doInitialReset').setValue(false)
      newSs.getSheetByName('⚙️Meta').getRange('B2:C4').setValues(this.ss.getSheetByName('⚙️Meta').getRange('B2:C4').getValues())

      DriveApp.getFileById(this.ss.getId()).setTrashed(true)
      _openLink_(newSs.getUrl(), 'New Client File Created')
    } catch (e) {
      if (newFile) newFile.setTrashed(true)
      _alert_('Some error occured!')
    }
  }

  this.replaceVersion = (str, newVersion) => {
    return str.replace(/(.* V)\d+(\.\d+)?/, `$1${newVersion}`);
  }
}

const showPlanningSheet = () => new App().showPlanningSheet()
const showHomeSheet = () => new App().showHomeSheet()
const showGainsightSheet = () => new App().showGainsightSheet()

const resetDatabases = () => new App().resetDatabases()

//Report Generation
const generateReport = () => new App().generateReport()
const saveReportChanges = () => new App().saveReportChanges()
const saveSolutions = (data, exCorePillar) => new App().saveSolutions(data, exCorePillar)
const saveDbData = (data) => new App().saveDbData(data)

//Gainsight
const mergeGainsightWithDatabase = () => new App().mergeGainsightWithDatabase()

//set goals
const setGoalsEffectiveCommunication = () => new App().showGoalsPrompt('1-effective-communication')
const setGoalsEmployeeExperience = () => new App().showGoalsPrompt('2-engaging-culture-experience')
const setGoalsEmployeeProductivity = () => new App().showGoalsPrompt('3-maximize-employee-productivity')

//set solutions
const setSolutionsEffectiveCommunication = () => new App().showSolutionsPrompt('1-effective-communication')
const setSolutionsEmployeeExperience = () => new App().showSolutionsPrompt('2-engaging-culture-experience')
const setSolutionsEmployeeProductivity = () => new App().showSolutionsPrompt('3-maximize-employee-productivity')

//set objectives
const setObjectivesEffectiveCommunication = () => new App().showObjectivesPrompt('1-effective-communication')
const setObjectivesEmployeeExperience = () => new App().showObjectivesPrompt('2-engaging-culture-experience')
const setObjectivesEmployeeProductivity = () => new App().showObjectivesPrompt('3-maximize-employee-productivity')

//export
const exportAsExcel = () => new App().exportSheet('xlsx')
const exportAsPDF = () => new App().exportSheet('pdf')

const checkMigration = () => new App().checkMigration()

const onOpen = () => _createMenu_('⚙️Success Planning', [{ caption: 'Export As Excel', action: 'exportAsExcel' },
{ caption: 'Export As PDF', action: 'exportAsPDF' }, null,
{ caption: 'Check For New Version', action: 'checkMigration' }])
//ARCHIVE - learn more//
const learnMoreEffectiveCommunication = () => _openLink_(`https://www.youtube.com/`)
const learnMoreEmployeeExperience = () => _openLink_(`https://www.youtube.com/`)
const learnMoreEmployeeProductivity = () => _openLink_(`https://www.youtube.com/`)

