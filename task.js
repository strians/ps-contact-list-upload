const async = require("async");

class Task {
  constructor(sharepoint) {
    this.sharepoint = sharepoint;
  }

  run() {
    return async.eachSeries(["Individual", "Group", "Medicare"], async list => this.processList(list));
  }

  processList(list) {
    return this.sharepoint.getContents(`/${list}`)
      .then((contents) => {
        let folders = contents.filter(item => item.__metadata.type === "SP.Folder");
        // For each folder:
        //   Check if contact log exists in Contact Log folder
        //     If it exists, pass
        //     Else upload copy of contact log template
        return Promise.resolve();
      });
  }
}

module.exports = Task;
