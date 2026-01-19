/**
 * RST Word Add-in Help Page JavaScript
 * Handles tab navigation and UI interactions
 */

(function () {
  'use strict';

  // DOM Elements
  const navTabs = document.querySelectorAll('.nav-tab');
  const sections = document.querySelectorAll('.help-section');
  const closeBtn = document.getElementById('close-help');
  const backToTopLink = document.getElementById('back-to-top');

  /**
   * Switch to a specific help section
   * @param {string} sectionId - The ID of the section to show
   */
  function showSection(sectionId) {
    // Update tabs
    navTabs.forEach(tab => {
      if (tab.dataset.section === sectionId) {
        tab.classList.add('active');
      } else {
        tab.classList.remove('active');
      }
    });

    // Update sections
    sections.forEach(section => {
      if (section.id === sectionId) {
        section.classList.add('active');
      } else {
        section.classList.remove('active');
      }
    });

    // Scroll to top of content
    const helpContent = document.querySelector('.help-content');
    if (helpContent) {
      helpContent.scrollTop = 0;
    }

    // Save last viewed section
    try {
      localStorage.setItem('rst-help-section', sectionId);
    } catch (e) {
      // localStorage may not be available
    }
  }

  /**
   * Initialize tab navigation
   */
  function initNavigation() {
    navTabs.forEach(tab => {
      tab.addEventListener('click', () => {
        const sectionId = tab.dataset.section;
        if (sectionId) {
          showSection(sectionId);
        }
      });
    });
  }

  /**
   * Initialize close button
   */
  function initCloseButton() {
    if (closeBtn) {
      closeBtn.addEventListener('click', () => {
        // Send message to parent to close help
        if (window.parent && window.parent !== window) {
          window.parent.postMessage({ action: 'closeHelp' }, '*');
        }
        // Fallback: try to close if opened as popup
        if (window.opener) {
          window.close();
        }
      });
    }
  }

  /**
   * Initialize back to top link
   */
  function initBackToTop() {
    if (backToTopLink) {
      backToTopLink.addEventListener('click', (e) => {
        e.preventDefault();
        const helpContent = document.querySelector('.help-content');
        if (helpContent) {
          helpContent.scrollTop = 0;
        }
        window.scrollTo({ top: 0, behavior: 'smooth' });
      });
    }
  }

  /**
   * Restore last viewed section from localStorage
   */
  function restoreLastSection() {
    try {
      const lastSection = localStorage.getItem('rst-help-section');
      if (lastSection) {
        const section = document.getElementById(lastSection);
        if (section) {
          showSection(lastSection);
        }
      }
    } catch (e) {
      // localStorage may not be available
    }
  }

  /**
   * Handle keyboard navigation
   */
  function initKeyboardNavigation() {
    document.addEventListener('keydown', (e) => {
      // Escape key closes help
      if (e.key === 'Escape') {
        if (closeBtn) {
          closeBtn.click();
        }
      }

      // Arrow keys for tab navigation
      if (e.key === 'ArrowLeft' || e.key === 'ArrowRight') {
        const activeTab = document.querySelector('.nav-tab.active');
        if (activeTab) {
          const tabs = Array.from(navTabs);
          const currentIndex = tabs.indexOf(activeTab);
          let newIndex;

          if (e.key === 'ArrowLeft') {
            newIndex = currentIndex > 0 ? currentIndex - 1 : tabs.length - 1;
          } else {
            newIndex = currentIndex < tabs.length - 1 ? currentIndex + 1 : 0;
          }

          const newTab = tabs[newIndex];
          if (newTab) {
            showSection(newTab.dataset.section);
            newTab.focus();
          }
        }
      }
    });
  }

  /**
   * Handle messages from parent window
   */
  function initMessageHandler() {
    window.addEventListener('message', (event) => {
      const data = event.data;

      if (data && data.action === 'showSection' && data.sectionId) {
        showSection(data.sectionId);
      }
    });
  }

  /**
   * Initialize the help page
   */
  function init() {
    initNavigation();
    initCloseButton();
    initBackToTop();
    initKeyboardNavigation();
    initMessageHandler();
    restoreLastSection();
  }

  // Initialize when DOM is ready
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
