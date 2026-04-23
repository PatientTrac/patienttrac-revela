#!/usr/bin/env node
/**
 * Batch VB6 Parser — processes all .frm files in a folder
 * Run: node batch-parse.js ./vb6-forms/
 * Outputs: migration-manifest.json + individual .parsed.json files
 */
const fs = require('fs')
const path = require('path')

const folder = process.argv[2] || './vb6-forms'

if (!fs.existsSync(folder)) {
  console.error(`Folder not found: ${folder}`)
  process.exit(1)
}

const forms = fs.readdirSync(folder).filter(f => f.endsWith('.frm'))
console.log(`Found ${forms.length} .frm files in ${folder}\n`)

const manifest = {
  generatedAt: new Date().toISOString(),
  totalForms: forms.length,
  forms: [],
  allValidationRules: [],
  allBusinessRules: [],
  allCalculations: [],
  allSQLQueries: [],
  suggestedJiraStories: [],
}

forms.forEach(formFile => {
  const filePath = path.join(folder, formFile)
  // Basic parse (import parse-vb6 logic inline for batch)
  const content = fs.readFileSync(filePath, 'utf-8')
  const lines = content.split('\n').map(l => l.trim())
  const formName = path.basename(formFile, '.frm')

  const summary = {
    formName,
    file: formFile,
    controlCount: (content.match(/^Begin\s+/gm) || []).length,
    hasValidation: /MsgBox|Validate|IsNumeric|IsDate/i.test(content),
    hasCalculations: /Total|Calc|Sum/i.test(content),
    hasSQL: /SELECT.*FROM/i.test(content),
    lineCount: lines.length,
    suggestedModule: suggestModule(formName),
    jiraEpic: suggestEpic(formName),
  }

  manifest.forms.push(summary)

  // Suggested Jira story per form
  manifest.suggestedJiraStories.push({
    summary: `Migrate VB6 ${formName} — extract & implement in React`,
    description: `Convert VB6 form ${formFile} to React component.\n` +
      `Controls: ~${summary.controlCount} | Has validation: ${summary.hasValidation} | Has SQL: ${summary.hasSQL}\n` +
      `Target module: ${summary.suggestedModule}\nEpic: ${summary.jiraEpic}`,
    issueType: 'Story',
    epic: summary.jiraEpic,
  })
})

// Write manifest
fs.writeFileSync('./migration-manifest.json', JSON.stringify(manifest, null, 2))

console.log('Migration Manifest Summary')
console.log('==========================')
manifest.forms.forEach(f => {
  console.log(`${f.formName.padEnd(30)} → ${f.suggestedModule.padEnd(20)} [${f.jiraEpic}]`)
})
console.log(`\nManifest saved: migration-manifest.json`)
console.log(`Jira stories to create: ${manifest.suggestedJiraStories.length}`)

function suggestModule(name) {
  const n = name.toLowerCase()
  if (/patient|intake|consult|history|medical/i.test(n)) return 'intake'
  if (/or|schedule|surg|operat|room/i.test(n)) return 'or-scheduling'
  if (/proposal|cost|fee|price|financ|billing/i.test(n)) return 'proposals'
  if (/inventory|supply|stock|implant|order/i.test(n)) return 'inventory'
  if (/account|revenue|payment|invoice|ledger/i.test(n)) return 'accounting'
  if (/staff|employee|doctor|nurse|credential/i.test(n)) return 'staff'
  return 'core'
}

function suggestEpic(name) {
  const mod = suggestModule(name)
  const map = {
    'intake': 'SCRUM-6',
    'or-scheduling': 'SCRUM-7',
    'proposals': 'SCRUM-8',
    'inventory': 'SCRUM-9',
    'accounting': 'SCRUM-10',
    'staff': 'SCRUM-11',
    'core': 'SCRUM-12',
  }
  return map[mod]
}
