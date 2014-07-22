/**
 * RoboScheduler 3000
 *
 * Schedule your meetings, automatically!
 * "Fine! I will make my OWN theme park!"
 *
 *                      .-.
 *                     (   )
 *                      '-'
 *                      J L
 *                      | |
 *                     J   L
 *                     |   |
 *                    J     L
 *                  .-'.___.'-.
 *                 /___________\
 *            _.-""'           `bmw._
 *          .'                       `.
 *        J                            `.
 *       F                               L
 *      J                                 J
 *     J                                  `
 *     |                                   L
 *     |                                   |
 *     |                                   |
 *     |                                   J
 *     |                                    L
 *     |                                    |
 *     |             ,.___          ___....--._
 *     |           ,'     `""""""""'           `-._
 *     |          J           _____________________`-.
 *     |         F         .-'   `-88888-'    `Y8888b.`.
 *     |         |       .'         `P'         `88888b \
 *     |         |      J       #     L      #    q8888b L
 *     |         |      |             |           )8888D )
 *     |         J      \             J           d8888P P
 *     |          L      `.         .b.         ,88888P /
 *     |           `.      `-.___,o88888o.___,o88888P'.'
 *     |             `-.__________________________..-'
 *     |                                    |
 *     |         .-----.........____________J
 *     |       .' |       |      |       |
 *     |      J---|-----..|...___|_______|
 *     |      |   |       |      |       |
 *     |      Y---|-----..|...___|_______|
 *     |       `. |       |      |       |
 *     |         `'-------:....__|______.J
 *     |                                  |
 *      L___                              |
 *          """----...______________....--'
 */

/**
 * Add a custom menu to the active spreadsheet.
 * @return {void}
 */
function onOpen() {
  SpreadsheetApp.getActiveSpreadsheet().addMenu("Periodic Scheduler", [
    {name: "Start Job ", functionName: "installTrigger_"},
    {name: "Stop Job", functionName: "removeTrigger_"},
    {name: "Run Manually", functionName: "run_"}
  ]);
  return true;
}
/**
 * Installs the trigger
 * @return {void}
 */
function installTrigger_() {
	RoboScheduler.installTrigger();
}

/**
 * Removes the trigger
 * @return {void}
 */
function removeTrigger_() {
	RoboScheduler.removeTrigger();
}

/**
 * Runs RoboScheduler
 * @return {void}
 */
function run_() {
	RoboScheduler.run();
}