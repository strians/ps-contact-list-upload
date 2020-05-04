const async = require("async");
const fs = require("fs");

class Task {
  constructor(sharepoint) {
    this.sharepoint = sharepoint;
  }

  run() {
    return async.eachSeries(["Individual", "Group", "Medicare"], async list => this.processList(list));
  }

  processList(list) {
    return this.sharepoint.getContents(`/${list}`)
      .then(contents => {
        let folders = contents.filter(item => {
          return item.__metadata.type === "SP.Folder" && item.Name !== "Forms"
        });
        return async.eachSeries(folders, async folder => this.processFolder(list, folder));
      });
  }

  processFolder(list, folder) {
    const documentName = `/${list}/${folder.Name}/Contact Log/Contact Log.xlsx`;
    return this.sharepoint.getDocument(documentName).then(contactLog => {
      console.log(`${documentName} exists`);
      return Promise.resolve();
    }).catch(err => {
      return this.uploadContactLogTemplate(`/${list}/${folder.Name}/Contact Log`);
    });
  }

  getBinaryData(filePath) {
    const base64 = fs.readFileSync(filePath, { encoding: 'base64' });
    const encoded = base64.replace(/^data:+[a-z]+\/+[a-z]+;base64,/, '');
    return Buffer.from(encoded, 'base64');
  }

  async uploadContactLogTemplate(path) {
    const fileName = "Contact Log.xlsx";
    const data = this.getBinaryData("./Contact Log.xlsx");
    try {
      const file = await this.sharepoint.createFile({ path, fileName, data });
      console.log(`Uploaded contact log to ${path}`);
    } catch (err) {
      console.error(`ERROR: Error uploading contact log to ${path}`);
    }
    return Promise.resolve();
  }
}

module.exports = Task;
