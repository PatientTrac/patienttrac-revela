#!/usr/bin/env node
/**
 * VB6 Form Parser for PatientTrac Revela
 * Run: node parse-vb6.js path/to/Form.frm
 * Outputs: JSON with fields, validation rules, events, and business logic
 */
const fs = require('fs')
const path = require('path')

function parseVB6Form(filePath) {
  const content = fs.readFileSync(filePath, 'utf-8')
  const lines = content.split('\n').map(l => l.trim())
  const result = {
    formName: path.basename(filePath, '.frm'),
    controls: [],
    variables: [],
    validationRules: [],
    businessRules: [],
    calculations: [],
    events: [],
    constants: [],
    sqlQueries: [],
  }

  let currentControl = null
  let inSub = false
  let currentSub = null
  let subLines = []

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i]

    // Constants
    if (/^Const\s+(\w+)\s*=\s*(.+)/i.test(line)) {
      const [, name, value] = line.match(/^Const\s+(\w+)\s*=\s*(.+)/i)
      result.constants.push({ name, value: value.trim() })
    }

    // Module-level variables
    if (/^(Dim|Public|Private)\s+(\w+)\s+As\s+(\w+)/i.test(line)) {
      const [, scope, name, type] = line.match(/^(Dim|Public|Private)\s+(\w+)\s+As\s+(\w+)/i)
      result.variables.push({ scope, name, type })
    }

    // Controls
    if (/^Begin\s+\w+\s+(\w+)/i.test(line)) {
      const [, name] = line.match(/^Begin\s+\w+\s+(\w+)/i)
      const type = line.match(/^Begin\s+(\w+)\s+/i)?.[1]
      currentControl = { name, type, properties: {} }
    }
    if (line === 'End' && currentControl) {
      result.controls.push(currentControl)
      currentControl = null
    }
    if (currentControl && /^\w+\s*=\s*.+/.test(line)) {
      const [prop, ...rest] = line.split('=')
      currentControl.properties[prop.trim()] = rest.join('=').trim()
    }

    // Subroutines and functions
    if (/^(Private\s+|Public\s+)?(Sub|Function)\s+(\w+)/i.test(line)) {
      inSub = true
      const match = line.match(/(?:Sub|Function)\s+(\w+)/i)
      currentSub = { name: match?.[1], lines: [], type: 'sub' }
      subLines = []
    }
    if (inSub && /^End\s+(Sub|Function)/i.test(line)) {
      if (currentSub) {
        currentSub.lines = subLines
        categorizeSubroutine(currentSub, result)
      }
      inSub = false
      currentSub = null
      subLines = []
    }
    if (inSub && currentSub) subLines.push(line)

    // SQL queries
    const sqlMatch = line.match(/["']([^"']*SELECT[^"']*FROM[^"']*)['"]/i)
    if (sqlMatch) result.sqlQueries.push({ foundIn: currentSub?.name, sql: sqlMatch[1] })
  }

  return result
}

function categorizeSubroutine(sub, result) {
  const code = sub.lines.join('\n').toLowerCase()
  const entry = { name: sub.name, logic: sub.lines.filter(l => l.length > 0) }

  // Validation patterns
  if (code.includes('msgbox') || code.includes('validate') || code.includes('if len(') ||
      code.includes('isnumeric') || code.includes('isdate') || code.includes('exit sub')) {
    result.validationRules.push(entry)
    return
  }
  // Calculation patterns
  if (code.includes('total') || code.includes('calc') || code.includes('sum') ||
      code.includes('= ') && (code.includes('+ ') || code.includes('* ') || code.includes('/ '))) {
    result.calculations.push(entry)
    return
  }
  // Event handlers
  if (sub.name.includes('_Click') || sub.name.includes('_Change') ||
      sub.name.includes('_Load') || sub.name.includes('_KeyPress')) {
    result.events.push(entry)
    return
  }
  // Everything else = business rules
  result.businessRules.push(entry)
}

function generateTypeScript(parsed) {
  const lines = []
  lines.push(`// Auto-generated from VB6: ${parsed.formName}`)
  lines.push(`// PatientTrac Revela — src/lib/rules/${toCamelCase(parsed.formName)}.ts`)
  lines.push(`// Review all generated code before use in production\n`)

  // Types from controls
  lines.push(`export interface ${parsed.formName}FormData {`)
  parsed.controls
    .filter(c => ['TextBox','ComboBox','ListBox','CheckBox','OptionButton'].includes(c.type))
    .forEach(c => {
      const tsType = c.type === 'CheckBox' ? 'boolean' : 'string'
      const label = c.properties.Caption || c.properties.Text || ''
      if (label) lines.push(`  // ${label.replace(/"/g,'')}`)
      lines.push(`  ${toCamelCase(c.name)}: ${tsType}`)
    })
  lines.push(`}\n`)

  // Constants
  if (parsed.constants.length > 0) {
    lines.push(`// Constants from VB6`)
    parsed.constants.forEach(c => {
      lines.push(`export const ${c.name} = ${c.value}`)
    })
    lines.push('')
  }

  // Validation stub
  if (parsed.validationRules.length > 0) {
    lines.push(`// Validation rules extracted from VB6`)
    lines.push(`export function validate${parsed.formName}(data: ${parsed.formName}FormData): string[] {`)
    lines.push(`  const errors: string[] = []\n`)
    parsed.validationRules.forEach(rule => {
      lines.push(`  // --- ${rule.name} ---`)
      rule.logic.slice(0, 5).forEach(l => lines.push(`  // VB6: ${l}`))
      lines.push(`  // TODO: Convert above VB6 logic to TypeScript\n`)
    })
    lines.push(`  return errors`)
    lines.push(`}\n`)
  }

  // Business rule stubs
  if (parsed.businessRules.length > 0) {
    lines.push(`// Business rules from VB6`)
    parsed.businessRules.slice(0, 5).forEach(rule => {
      lines.push(`export function ${toCamelCase(rule.name)}(data: ${parsed.formName}FormData) {`)
      rule.logic.slice(0, 3).forEach(l => lines.push(`  // VB6: ${l}`))
      lines.push(`  // TODO: implement`)
      lines.push(`}\n`)
    })
  }

  return lines.join('\n')
}

function toCamelCase(str) {
  return str.replace(/[_\s](.)/g, (_, c) => c.toUpperCase())
            .replace(/^(.)/, c => c.toLowerCase())
}

// CLI entrypoint
const args = process.argv.slice(2)
if (args.length === 0) {
  console.log('Usage: node parse-vb6.js <path/to/Form.frm> [--ts]')
  console.log('  --ts    Also generate TypeScript output file')
  process.exit(0)
}

const filePath = args[0]
const generateTS = args.includes('--ts')

if (!fs.existsSync(filePath)) {
  console.error(`File not found: ${filePath}`)
  process.exit(1)
}

const parsed = parseVB6Form(filePath)

// JSON output
const jsonOut = filePath.replace('.frm', '.parsed.json')
fs.writeFileSync(jsonOut, JSON.stringify(parsed, null, 2))
console.log(`\nParsed: ${parsed.formName}`)
console.log(`  Controls:         ${parsed.controls.length}`)
console.log(`  Variables:        ${parsed.variables.length}`)
console.log(`  Validation rules: ${parsed.validationRules.length}`)
console.log(`  Business rules:   ${parsed.businessRules.length}`)
console.log(`  Calculations:     ${parsed.calculations.length}`)
console.log(`  Events:           ${parsed.events.length}`)
console.log(`  SQL queries:      ${parsed.sqlQueries.length}`)
console.log(`\nJSON saved: ${jsonOut}`)

if (generateTS) {
  const ts = generateTypeScript(parsed)
  const tsOut = filePath.replace('.frm', '.generated.ts')
  fs.writeFileSync(tsOut, ts)
  console.log(`TypeScript saved: ${tsOut}`)
}
