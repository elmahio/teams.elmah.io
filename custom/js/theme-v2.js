/**
 * Class for managing Microsoft Teams v2 themes
 * idea borrowed from the Dizz: https://github.com/richdizz/Microsoft-Teams-Tab-Themes/blob/master/app/config.html
 * Updated on 03/04/2024
 */
var TeamsTheme = (function () {
    function TeamsTheme() {
    }
    /**
     * Set up themes on a page
     */
    TeamsTheme.fix = function (context) {
        microsoftTeams.app.initialize().then(() => {
            microsoftTeams.registerOnThemeChangeHandler(TeamsTheme.themeChanged);
            if (context) {
                TeamsTheme.themeChanged(context.app.theme);
            } else {
                microsoftTeams.app.getContext().then(function (context) {
                    TeamsTheme.themeChanged(context.app.theme);
                });
            }
        });
    };
    /**
     * Manages theme changes
     * @param theme default|contrast|dark
     */
    TeamsTheme.themeChanged = function (theme) {
        const bodyElement = document.querySelector('body');
        switch (theme) {
            case "dark":
            case "contrast":
                bodyElement.className = "theme-" + theme;
                break;
            case "default":
                bodyElement.className = "";
        }
    };
    return TeamsTheme;
}());

