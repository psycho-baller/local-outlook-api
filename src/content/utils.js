/**
 * Utility functions for Outlook Web Client interaction
 */

/**
 * Waits for an element to appear in the DOM
 * @param {string} selector - CSS selector for the element
 * @param {number} timeout - Timeout in milliseconds
 * @returns {Promise<Element>} The found element
 */
function waitForElement(selector, timeout = 10000) {
  return new Promise((resolve, reject) => {
    const startTime = Date.now();

    const checkElement = () => {
      const element = document.querySelector(selector);

      if (element) {
        resolve(element);
        return;
      }

      if (Date.now() - startTime > timeout) {
        reject(new Error(`Timeout waiting for element: ${selector}`));
        return;
      }

      setTimeout(checkElement, 100);
    };

    checkElement();
  });
}

/**
 * Sleep function for waiting
 * @param {number} ms - Milliseconds to sleep
 * @returns {Promise<void>}
 */
function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

// Export the utility functions
export { waitForElement, sleep };
