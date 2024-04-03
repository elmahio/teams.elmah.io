microsoftTeams.app.initialize().then(() => {
    microsoftTeams.app.getContext().then(function (context) {
        TeamsTheme.fix(context);
    });

    microsoftTeams.pages.getConfig().then(function (settings) {
        document.querySelector("#webhook").value = settings.webhookUrl;
        microsoftTeams.pages.config.setValidityState(true);
    });

    microsoftTeams.pages.config.registerOnSaveHandler((saveEvent) => {
        const configPromise = microsoftTeams.pages.config.setConfig({
            entityId: "myconfig",
            contentUrl: "https://dev.seomat.com/teams/install.html",
            configName: "myconfig"
        });

        configPromise.
            then((result) => {saveEvent.notifySuccess()}).
            catch((error) => {saveEvent.notifyFailure("failure message")});
    });
});

function copyWebhookToClipborad() {
    var copyText = document.getElementById("webhook");
    copyText.select();
    document.execCommand("copy");
    if (copyText.value.length <= 0) {
       alert("Some Error");
    }
}