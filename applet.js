const Lang = imports.lang
const Soup = imports.gi.Soup
const Util = imports.misc.util
const Applet = imports.ui.applet
const Mainloop = imports.mainloop
const Settings = imports.ui.settings

const DEFAULT_ICON = 'o365.svg'
const NOTIFICATION_ICON = 'o365-notification.svg'
const ERROR_ICON = 'error.svg'

function main(metadata, orientation, panel_height, instanceId) {
  return new MyApplet(metadata, instanceId)
}

function MyApplet(metadata, instanceId) {
  this._init(metadata, instanceId)
}

MyApplet.prototype = {
    
    __proto__: Applet.IconApplet.prototype,

    _init: function(metadata, instanceId) {
      Applet.IconApplet.prototype._init.call(this, null, 30, instanceId)
      this.settings = new Settings.AppletSettings(this, metadata.uuid, instanceId)
      this.currentDir = imports.ui.appletManager.appletMeta[metadata.uuid].path
      this.set_applet_tooltip('Click here to go to your Office 365 Outlook inbox')
      for (setting of ['email', 'password', 'refreshTime']) {
        this.settings.bindProperty(Settings.BindingDirection.IN, setting, setting, this._update)
      }
      this._set_icon(DEFAULT_ICON)
      this._mainLoop()
    },

    _update: function() {
      let httpSession = new Soup.Session()
      let authUri = new Soup.URI('https://outlook.office365.com/api/v1.0/Me/Folders/Inbox')
      authUri.set_user(this.email)
      authUri.set_password(this.password)
      let req = new Soup.Message({method: 'GET', uri: authUri})
      httpSession.queue_message(req, (_, res) => {
        let iconName = ERROR_ICON
        if (res.status_code == 200) {
          let data = JSON.parse(res.response_body.data)
          if (data.UnreadItemCount > 0) {
            iconName = NOTIFICATION_ICON
          } else {
            iconName = DEFAULT_ICON
          }
        }
        this._set_icon(iconName)
      })
    },

    _mainLoop: function () {
      this._update()
      this.update_timeout_id = Mainloop.timeout_add(this.refreshTime * 1000, Lang.bind(this, this._mainLoop))
    },

    on_applet_clicked: function() {
      Util.spawnCommandLine("xdg-open https://outlook.office365.com/owa/")
    },

    on_applet_removed_from_panel: function() {
      if (this.update_timeout_id > 0) {
          Mainloop.source_remove(this.update_timeout_id)
          this.update_timeout_id = 0
      }
    },

    _set_icon: function(iconName) {
      this.set_applet_icon_path(this.currentDir + '/icons/' + iconName)
    }
}
