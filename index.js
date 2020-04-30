const SharePoint = require("@paulholden/sharepoint");
const rc = require('rc');
const Task = require('./task.js');

const APP = "pscontact";
const conf = rc(APP);

if (!conf.configs) {
  console.error(`\
No config files were found. Configuration is loaded from:
  1) Any file passed via --config argument
  2) Any .${APP}rc file found in local or parent directories
     * Note: not available for packaged binaries
  3) $HOME/.${APP}rc
  4) $HOME/.${APP}/config
  5) $HOME/.config/${APP}
  6) $HOME/.config/${APP}/config
  7) /etc/${APP}rc
  8) /etc/${APP}/config

Configurations are loaded in JSON or INI format.
Data is merged down; earlier configs override those that follow.

Example configuration (JSON):
   {
     "sharepoint": {
       "url": "https://my.sharepoint.com/sites/MySite",
       "username": "johndoe",
       "password": "abcd1234"
     }
   }

About configuration settings:
  - sharepoint.url: The URL for your SharePoint site
  - sharepoint.username: The username used to access the SharePoint site
  - sharepoint.password: The password used to access the SharePoint site
` );
  return;
}

let sharepoint = new SharePoint(conf.sharepoint.url);

sharepoint.authenticate(conf.sharepoint.username, conf.sharepoint.password)
  .then(() => {
    return sharepoint.getWebEndpoint();
  })
  .then(() => {
    let task = new Task(sharepoint);
    return task.run();
  })
  .catch(err => {
    console.error(err);
  });
