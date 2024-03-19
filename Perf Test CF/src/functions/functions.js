/* global clearInterval, console, setInterval */

/**
 * Add two numbers
 * @customfunction
 * @param {number} first First number
 * @param {number} second Second number
 * @returns {number} The sum of the two numbers.
 */
export function add(first, second) {
  return first + second;
}

/**
 * Displays the current time once a second
 * @customfunction
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 */
export function clock(invocation) {
  const timer = setInterval(() => {
    const time = currentTime();
    invocation.setResult(time);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Returns the current time
 * @returns {string} String with the current time formatted for the current locale.
 */
export function currentTime() {
  return new Date().toLocaleTimeString();
}

/**
 * Increments a value once a second.
 * @customfunction
 * @param {number} incrementBy Amount to increment
 * @param {CustomFunctions.StreamingInvocation<number>} invocation
 */
export function increment(incrementBy, invocation) {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    invocation.setResult(result);
  }, 1000 * 10);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Writes a message to console.log().
 * @customfunction LOG
 * @param {string} message String to write.
 * @returns String to write.
 */
export function logMessage(message) {
  console.log(message);

  return message;
}

// Write my own functions starting here

/**
 * Return random int - 10
 * @customfunction
 * @returns {number} Return random int - 10.
 */
function returnInt() {
  var timestamp = new Date().toISOString();
  console.log(`[${timestamp}] `, 'Call returnInt.');
  return Math.floor(Math.random() * 10);
}

/**
 * Return random int - 100
 * @customfunction
 * @returns {number} Return random int - 100.
 */
function returnIntPromise() {
  var timestamp = new Date().toISOString();
  console.log(`[${timestamp}] `, 'Call returnIntPromise.');
  return new Promise(function (resolve) {
      setTimeout(function () {
          resolve(Math.floor(Math.random() * 100));
      }, 1000);
  });
}

/**
 * Return random int - 1000
 * @customfunction
 * @volatile
 * @returns {number} Return random int - 1000.
 */
function returnIntVolatile() {
  var timestamp = new Date().toISOString();
  console.log(`[${timestamp}] `, 'Call returnIntVolatile.');
  return Math.floor(Math.random() * 1000)
}

/**
 * Return current time
 * @customfunction
 * @volatile
 * @returns {string} Return current time.
 */
function returnStringVolatile() {
  var timestamp = new Date().toISOString();
  console.log(`[${timestamp}] `, 'Call returnStringVolatile.');
  return currentTime();
}

/**
 * Return current time
 * @customfunction
 * @param {CustomFunctions.StreamingInvocation<number>} invocation
 */
export function returnStringStream(invocation) {
  let result = currentTime();
  invocation.setResult(result);
  var timestamp = new Date().toISOString();
  console.log(`[${timestamp}] `, 'Call returnStringStream#1.');
  const timer = setInterval(() => {
    //var timestamp = new Date().toISOString();
    result = currentTime();
    invocation.setResult(result);
    console.log(`[${result}] `, 'Call returnStringStream#2.');
  }, 1000 * 60 * 5);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Return current time
 * @customfunction
 * @returns {string[][]} Return current time.
 */
export function returnStringDynamicArray() {
  var timestamp = new Date().toISOString();
  console.log(`[${timestamp}] `, 'Call returnStringDynamicArray.');
  return [[currentTime(), currentTime(), currentTime()]];
}

/**
 * Return nested range
 * @customfunction
 * @param {string[][]} values Multiple ranges of values.
 * @returns {string} Return nested range.
 */
export function returnStringNested(range) {
  var cell = range[0];
  var timestamp = new Date().toISOString();
  console.log(`[${timestamp}] `, 'Call returnStringNested.');
  return cell;
}

/**
 * Wait before returning current time
 * @customfunction
 * @returns {string} Return current time.
 */
export function returnStringWait() {
  var timestamp = new Date().toISOString();
  console.log(`[${timestamp}] `, 'Call returnStringWait.');
  var num = 0;
  for (var i = 0; i < 100000000000; i++) {
    num++;
  }
  return timestamp;
}

/**
 * Wait before returning current time
 * @customfunction
 * @volatile
 * @returns {string} Return current time.
 */
export function returnStringWaitVolatile() {
  var timestamp = new Date().toISOString();
  console.log(`[${timestamp}]`, 'Call returnStringWaitVolatile.');
  var num = 0;
  for (var i = 0; i < 100000000000; i++) {
    num++;
  }
  return timestamp;
}

/**
 * Input a string to output
 * @customfunction
 * @returns {string} Return input.
 */
export function inputString(str) {
  var timestamp = new Date().toISOString();
  console.log(`[${timestamp}]`, 'Call inputString.');
  return str;
}