const path = require("path")
const fs = require("fs")
const unzipper = require("unzipper")
const { resolve } = require("path")
const { readdir } = require("fs").promises

const getUrls = require("get-urls")

async function getFiles(dir) {
  const dirents = await readdir(dir, { withFileTypes: true })
  const files = await Promise.all(
    dirents.map((dirent) => {
      const res = resolve(dir, dirent.name)
      return dirent.isDirectory() ? getFiles(res) : res
    })
  )
  return Array.prototype.concat(...files)
}

let processOfficeFile = async (filePath) => {
  let allUrls = new Set()
  fs.rmdirSync(`${__dirname}/temp`, { recursive: true })
  fs.mkdirSync(`${__dirname}/temp`)
  await fs
    .createReadStream(filePath)
    .pipe(unzipper.Extract({ path: `${__dirname}/temp` }))
    .promise()
  let files = await getFiles(`${__dirname}/temp`)
  for (let i = 0; i < files.length; i++) {
    let file = files[i]
    if (file.includes(".xml") || file.includes(".rels")) {
      let contents = fs.readFileSync(file).toString()
      let urls = getUrls(contents)
      for (u of urls) {
        allUrls.add(u)
      }
    }
  }
  let arr = [...allUrls]
  arr = arr.filter(
    (x) =>
      !x.includes("schemas.openxml") &&
      !x.includes("vnd.openxmlformats") &&
      !x.includes("schemas.microsoft") &&
      !/http:\/\/image[0-9]*\./.test(x) &&
      !/w3\.org.*XMLSchema/.test(x)
  )
  return "\n\n===== " + path.parse(filePath).base + "\n" + arr.join("\n")
}
let processZip = async () => {
  fs.rmdirSync(`${__dirname}/temp2`, { recursive: true })
  fs.mkdirSync(`${__dirname}/temp2`)
  await fs
    .createReadStream(`${__dirname}/input.zip`)
    .pipe(unzipper.Extract({ path: `${__dirname}/temp2` }))
    .promise()
  let files = await getFiles(`${__dirname}/temp2`)
  let officeFiles = files.filter((f) => f.includes(".pptx") || f.includes(".docx"))
  let ret = ""
  for (let i = 0; i < officeFiles.length; i++) {
    ret += await processOfficeFile(officeFiles[i])
  }
  fs.writeFileSync(`${__dirname}/output.txt`, ret)
}

processZip()
