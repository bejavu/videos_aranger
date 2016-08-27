var fs = require('fs')
var xlsx = require('node-xlsx')
var commandLineArgs = require('command-line-args')
var mkdirp = require('mkdirp')
var async = require('async')

var optionDefinitions = [
  { name: 'excel_filename', alias: 'f', type: String },
  { name: 'input_folder', alias: 'i', type: String },
  { name: 'output_folder', alias: 'o', type: String },
  { name: 'file_ending', alias: 'e', type: String }
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

function copy_file(source, target, cb) {
  var cbCalled = false

  var rd = fs.createReadStream(source)
  rd.on("error", function(err) {
    done(err)
  })
  var wr = fs.createWriteStream(target)
  wr.on("error", function(err) {
    done(err)
  })
  wr.on("close", function(ex) {
    done()
  })
  rd.pipe(wr)

  function done(err) {
    if (!cbCalled) {
      cb(err)
      cbCalled = true
    }
  }
}

var move_files = function (options, callback) {
  var input_folder = options.input_folder || '.'
  var output_folder = options.output_folder || '.'
  var excel_filename = options.excel_filename
  var file_ending = options.file_ending || 'mp4'
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
      async.waterfall([
        function (cb) {
          copy_file(input_folder + '/names/' + name + '.' + file_ending, out_folder + '/' + name + '.' + file_ending, cb)
        },
        function (cb) {
          copy_file(input_folder + '/dates/' + date + '.' + file_ending, out_folder + '/' + date + '.' + file_ending, cb)
        },
        function (cb) {
          copy_file(input_folder + '/signs/' + sign + '.' + file_ending, out_folder + '/' + sign + '.' + file_ending, cb)
        }
      ], cb)
    })
  }, callback)
}

move_files(options, function (err) {
  if (err) return console.error('!!!', err)
  console.log('Done!')
  process.exit(0)
})