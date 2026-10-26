/* global CustomFunctions */

/**
 * Adds two numbers together.
 * @customfunction
 * @param {number} first First number
 * @param {number} second Second number
 * @returns {number} The sum of the two numbers.
 */
function add(first, second) {
  return (Number(first) || 0) + (Number(second) || 0);
}

/**
 * Multiplies two numbers.
 * @customfunction
 * @param {number} first First number
 * @param {number} second Second number
 * @returns {number} The product of the two numbers.
 */
function multiply(first, second) {
  return (Number(first) || 0) * (Number(second) || 0);
}

/**
 * Streams current time every intervalMs.
 * @customfunction
 * @param {number} [intervalMs] Interval in milliseconds (default: 1000)
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Streaming invocation object
 */
function streamTimer(intervalMs, invocation) {
  const delay = Math.max(250, Number(intervalMs) || 1000);
  
  // Send immediate initial value
  invocation.setResult(new Date().toLocaleTimeString());

  const timer = setInterval(() => {
    invocation.setResult(new Date().toLocaleTimeString());
  }, delay);

  // Invocation canceled hook when cell formula changes or is cleared
  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

// Associate with Custom Functions Runtime IDs
CustomFunctions.associate('ADD', add);
CustomFunctions.associate('MULTIPLY', multiply);
CustomFunctions.associate('STREAMTIMER', streamTimer);
