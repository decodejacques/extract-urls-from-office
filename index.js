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
var myArgs = process.argv.slice(2)
let main = async () => {
  let allUrls = new Set()
  fs.rmdirSync(`${__dirname}/temp`, { recursive: true })
  fs.mkdirSync(`${__dirname}/temp`)
  await fs
    .createReadStream(myArgs[0])
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
  console.log(arr.join("\n"))
}
main()
