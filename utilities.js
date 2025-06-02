/**
 * Validates an email address format
 * @param {string} email - The email address to validate
 * @returns {boolean} True if the email is valid
 */
function isValidEmail(email) {
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailRegex.test(email);
  }
  
  /**
   * Sanitizes input to prevent XSS and other security issues
   * @param {string} input - The input to sanitize
   * @returns {string} The sanitized input
   */
  function sanitizeInput(input) {
    if (typeof input !== 'string') {
      return input;
    }
    return input
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }