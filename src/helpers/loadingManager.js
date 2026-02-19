// helpers/loadingManager.js
import { CSS_CLASSES } from './constants.js';

class LoadingManager {
  constructor() {
    this.activeOperations = new Set();
  }

  showLoading(element, operation = 'default') {
    this.activeOperations.add(operation);
    
    if (element && element.length) {
      element.addClass(CSS_CLASSES.LOADING);
      
      // Add spinner overlay if not present
      if (!element.find('.sheexcel-spinner').length) {
        const spinner = $(`
          <div class="sheexcel-spinner">
            <div class="sheexcel-spinner-icon">
              <span class="sheexcel-spinner-symbol">⏳</span>
            </div>
            <div class="sheexcel-spinner-text">Loading...</div>
          </div>
        `);
        element.append(spinner);
      }
    }
  }

  hideLoading(element, operation = 'default') {
    this.activeOperations.delete(operation);
    
    if (element && element.length) {
      element.removeClass(CSS_CLASSES.LOADING);
      element.find('.sheexcel-spinner').remove();
    }
  }

  showError(element, message, temporary = true) {
    if (!element || !element.length) return;

    element.removeClass(CSS_CLASSES.LOADING);
    element.addClass(CSS_CLASSES.ERROR);
    
    const errorDiv = $(`
      <div class="sheexcel-error-message">
        <span class="sheexcel-error-icon">⚠</span>
        <span>${message}</span>
        <button class="sheexcel-error-dismiss" title="Dismiss">
          <span class="sheexcel-error-dismiss-icon">✕</span>
        </button>
      </div>
    `);
    
    // Remove existing error messages
    element.find('.sheexcel-error-message').remove();
    element.append(errorDiv);
    
    // Add dismiss handler
    errorDiv.find('.sheexcel-error-dismiss').on('click', () => {
      errorDiv.remove();
      element.removeClass(CSS_CLASSES.ERROR);
    });

    // Auto-dismiss after 5 seconds if temporary
    if (temporary) {
      setTimeout(() => {
        errorDiv.fadeOut(() => {
          errorDiv.remove();
          element.removeClass(CSS_CLASSES.ERROR);
        });
      }, 5000);
    }
  }

  clearError(element) {
    if (element && element.length) {
      element.removeClass(CSS_CLASSES.ERROR);
      element.find('.sheexcel-error-message').remove();
    }
  }

  async withLoading(element, operation, promise) {
    this.showLoading(element, operation);
    
    try {
      const result = await promise;
      this.hideLoading(element, operation);
      this.clearError(element);
      return result;
    } catch (error) {
      this.hideLoading(element, operation);
      this.showError(element, error.message);
      throw error;
    }
  }

  isLoading(operation = null) {
    if (operation) {
      return this.activeOperations.has(operation);
    }
    return this.activeOperations.size > 0;
  }
}

// Export singleton instance
export const loadingManager = new LoadingManager();