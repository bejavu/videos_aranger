var fs = require('fs')
var xlsx = require('node-xlsx')
var commandLineArgs = require('command-line-args')
var mkdirp = require('mkdirp')
var async = require('async')

var optionDefinitions = [
  { name: 'excel_filename', alias: 'f', type: String },
  { name: 'input_folder', alias: 'i', type: String },
  { name: 'output_folder', alias: 'o', type: String }
]

var options = commandLineArgs(optionDefinitions)

var read_excel = function (excel_filename) {
  var workSheets = xlsx.parse(excel_filename)
  return workSheets
}

if (!options.excel_filename) {
  console.error('No excel_filename')
  return
}

var create_folder = function (folder_name, callback) {
  mkdirp(folder_name, callback)
}

var copy_file = function (from, to) {
  console.log('copying', from, 'to', to)
  fs.createReadStream(from).pipe(fs.createWriteStream(to))
  console.log('done copying', from, 'to', to)
}

var move_files = function (options, callback) {
  var input_folder = options.input_folder || '.'
  var output_folder = options.output_folder || '.'
  var excel_filename = options.excel_filename
  if (!excel_filename) throw new Error('file not found')
  var workSheets = read_excel(options.excel_filename)
  console.log('excel data: ', JSON.stringify(workSheets))
  if (!workSheets || !workSheets.length || !workSheets[0].data) {
    throw new Error('problem with file')
  }
  async.each(workSheets[0].data, function (row, cb) {
    var name = row[0]
    var date = row[1]
    var sign = row[2]
    var out_folder = output_folder + '/' + name + '_' + date + '_' + sign
    create_folder(out_folder, function (err) {
      if (err) return cb(err)
      copy_file(input_folder + '/names/' + name + '.txt', out_folder + '/' + name + '.txt')
      copy_file(input_folder + '/dates/' + date + '.txt', out_folder + '/' + date + '.txt')
      copy_file(input_folder + '/signs/' + sign + '.txt', out_folder + '/' + sign + '.txt')
      cb()
    })
  }, callback)
}

move_files(options, function (err) {
  if (err) return console.error('!!!', err)
  console.log('Done!')
  process.exit(0)
})