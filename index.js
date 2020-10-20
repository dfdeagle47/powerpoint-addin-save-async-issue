/**
 * Logs a message in the logs "list".
 * @param {String} message - Message to log
 */
function logMessage(message) {
  const pElement = document.createElement('li')
  pElement.textContent = message

  document.querySelector('#logs').appendChild(pElement)
}

/**
 * Returns the value of settings saved in the document settings.
 * @param {String} name - Name of the setting
 */
function getState(name) {
  return Office.context.document.settings.get(name)
}

/**
 * Sets the value of a setting in the document settings.
 * @param {String} name - Name of the setting
 * @param {String} val - Value for this setting
 */
function setState(name, val) {
  logMessage(`Calling \`setState\` with ${name} and ${val}`)

  return Office.context.document.settings.set(name, val)
}

/**
 * Saves the document settings so the changes are persisted in the
 * document.
 */
function saveState() {
  logMessage('Calling `saveState`')

  return Office.context.document.settings.saveAsync((asyncResult) => {
    logMessage(
      `\`saveAsync\` callback called with ${JSON.stringify(asyncResult)}`
    )

    const newState = getState('state')
    logMessage(`newState ${newState}`)
    document.querySelector('#state').textContent = newState
  })
}

/**
 * Increments the counter by 1.
 */
function incrementState() {
  const currentState = getState('state')

  const newState = typeof currentState === 'number' ? currentState + 1 : 0

  setState('state', newState)

  saveState()
}

Office.onReady().then(() => {
  document
    .querySelector('#increment-button')
    .addEventListener('click', incrementState)

  const initialState = getState('state')

  // Initialise the state with the value in the document settings if
  // present.
  if (typeof initialState === 'number') {
    document.querySelector('#state').textContent = initialState
    logMessage(`initialState ${initialState}`)
  }
})
