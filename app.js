document.addEventListener('DOMContentLoaded', () => {
  const slides = document.querySelectorAll('.slide');
  const totalSlides = slides.length;
  let currentSlideIndex = 0;

  const currentSlideNumEl = document.getElementById('currentSlideNum');
  const totalSlidesNumEl = document.getElementById('totalSlidesNum');
  const progressBar = document.getElementById('progressBar');
  const currentSectionNameEl = document.getElementById('currentSectionName');
  
  const notesDrawer = document.getElementById('speakerNotesDrawer') || document.getElementById('notesDrawer');
  const notesTextContainer = document.getElementById('notesTextContainer') || document.getElementById('notesContent');
  const toggleNotesBtn = document.getElementById('toggleNotesBtn');
  const closeNotesBtn = document.getElementById('closeNotesBtn');
  
  const thumbnailsOverlay = document.getElementById('thumbnailsOverlay') || document.getElementById('thumbnailsModal');
  const toggleThumbnailsBtn = document.getElementById('toggleThumbnailsBtn');
  const closeThumbnailsBtn = document.getElementById('closeThumbnailsBtn');
  const thumbnailsGrid = document.getElementById('thumbnailsGrid');
  const toggleFullscreenBtn = document.getElementById('toggleFullscreenBtn');

  const nextSlideBtnEl = document.getElementById('nextSlideBtn');
  const prevSlideBtnEl = document.getElementById('prevSlideBtn');

  if (totalSlidesNumEl) totalSlidesNumEl.textContent = totalSlides;

  // Build Thumbnail Grid
  slides.forEach((slide, idx) => {
    const titleEl = slide.querySelector('.slide-title, .hero-title, .section-divider-title');
    const titleText = titleEl ? titleEl.textContent.trim() : `Slide ${idx + 1}`;
    
    const card = document.createElement('div');
    card.className = `thumb-card ${idx === 0 ? 'active' : ''}`;
    card.innerHTML = `
      <span class="thumb-number">SLIDE ${idx + 1}</span>
      <h4 class="thumb-title">${titleText}</h4>
    `;
    card.addEventListener('click', () => {
      goToSlide(idx);
      closeThumbnails();
    });
    thumbnailsGrid.appendChild(card);
  });

  function updateSlideView() {
    slides.forEach((slide, idx) => {
      if (idx === currentSlideIndex) {
        slide.classList.add('active');
      } else {
        slide.classList.remove('active');
      }
    });

    // Update Counter & Progress
    if (currentSlideNumEl) currentSlideNumEl.textContent = currentSlideIndex + 1;
    if (progressBar) {
      const progressPercent = ((currentSlideIndex + 1) / totalSlides) * 100;
      progressBar.style.width = `${progressPercent}%`;
    }

    // Update Section Name in Footer
    const activeSlide = slides[currentSlideIndex];
    const categoryEl = activeSlide.querySelector('.slide-category');
    if (currentSectionNameEl) {
      currentSectionNameEl.textContent = categoryEl ? categoryEl.textContent.trim() : 'ژیری دەستکرد';
    }

    // Update Speaker Notes
    const activeNotes = activeSlide.dataset.notes || 'هیچ تێبینییەک بۆ ئەم سلایدە تۆمار نەکراوە.';
    if (notesTextContainer) {
      notesTextContainer.innerHTML = activeNotes;
    }

    // Update Active Thumbnail
    const thumbCards = thumbnailsGrid.querySelectorAll('.thumb-card');
    thumbCards.forEach((tc, idx) => {
      if (idx === currentSlideIndex) {
        tc.classList.add('active');
      } else {
        tc.classList.remove('active');
      }
    });
  }

  function goToSlide(index) {
    if (index >= 0 && index < totalSlides) {
      currentSlideIndex = index;
      updateSlideView();
    }
  }

  function nextSlide() {
    if (currentSlideIndex < totalSlides - 1) {
      goToSlide(currentSlideIndex + 1);
    }
  }

  function prevSlide() {
    if (currentSlideIndex > 0) {
      goToSlide(currentSlideIndex - 1);
    }
  }

  function toggleNotes() {
    notesDrawer.classList.toggle('open');
  }

  function openThumbnails() {
    thumbnailsOverlay.classList.add('active');
  }

  function closeThumbnails() {
    thumbnailsOverlay.classList.remove('active');
  }

  function toggleFullscreen() {
    if (!document.fullscreenElement) {
      document.documentElement.requestFullscreen().catch(err => {
        console.error('Fullscreen error:', err);
      });
    } else {
      if (document.exitFullscreen) {
        document.exitFullscreen();
      }
    }
  }

  // Event Listeners
  if (nextSlideBtnEl) nextSlideBtnEl.addEventListener('click', nextSlide);
  if (prevSlideBtnEl) prevSlideBtnEl.addEventListener('click', prevSlide);
  if (toggleNotesBtn) toggleNotesBtn.addEventListener('click', toggleNotes);
  if (closeNotesBtn) closeNotesBtn.addEventListener('click', () => notesDrawer.classList.remove('open'));
  if (toggleThumbnailsBtn) toggleThumbnailsBtn.addEventListener('click', openThumbnails);
  if (closeThumbnailsBtn) closeThumbnailsBtn.addEventListener('click', closeThumbnails);
  if (toggleFullscreenBtn) toggleFullscreenBtn.addEventListener('click', toggleFullscreen);

  // Bulletproof Hardware Presenter Remote & Keyboard Controller
  document.addEventListener('keydown', (e) => {
    // If thumbnails active, handle Escape
    if (thumbnailsOverlay.classList.contains('active')) {
      if (e.key === 'Escape') closeThumbnails();
      return;
    }

    // Key Codes for Presenter Remotes (Logitech, Targus, Kensington, Generic Laser Clickers)
    const key = e.key;

    // NEXT SLIDE (Forward Keys sent by Remotes & Keyboard)
    if (
      key === 'PageDown' || 
      key === 'ArrowRight' || 
      key === 'ArrowDown' || 
      key === ' ' || 
      key === 'Enter' ||
      key === 'n' ||
      key === 'N'
    ) {
      e.preventDefault();
      nextSlide();
      return;
    }

    // PREVIOUS SLIDE (Backward Keys sent by Remotes & Keyboard)
    if (
      key === 'PageUp' || 
      key === 'ArrowLeft' || 
      key === 'ArrowUp' || 
      key === 'Backspace' ||
      key === 'p' ||
      key === 'P'
    ) {
      e.preventDefault();
      prevSlide();
      return;
    }

    // FULLSCREEN TOGGLE (F5, Shift+F5, F key)
    if (key === 'F5' || key === 'f' || key === 'F') {
      e.preventDefault();
      toggleFullscreen();
      return;
    }

    // THUMBNAILS OVERLAY (T, Escape)
    if (key === 't' || key === 'T' || key === 'Escape') {
      e.preventDefault();
      if (thumbnailsOverlay.classList.contains('active')) {
        closeThumbnails();
      } else {
        openThumbnails();
      }
      return;
    }

    // SPEAKER NOTES DRAWER TOGGLE (Alt+N or Shift+N or UI Button)
    if (e.altKey && (key === 'n' || key === 'N')) {
      e.preventDefault();
      toggleNotes();
      return;
    }
  });

  // Touch & Swipe Support
  let touchStartX = 0;
  document.addEventListener('touchstart', (e) => {
    touchStartX = e.changedTouches[0].screenX;
  });

  document.addEventListener('touchend', (e) => {
    const touchEndX = e.changedTouches[0].screenX;
    const diffX = touchEndX - touchStartX;
    if (Math.abs(diffX) > 50) {
      if (diffX < 0) {
        nextSlide();
      } else {
        prevSlide();
      }
    }
  });

  // Prompt Library Modal controls
  const openPromptModalBtn = document.getElementById('openPromptLibraryBtn');
  const closePromptModalBtn = document.getElementById('closePromptModalBtn');
  const promptModal = document.getElementById('promptModal');

  if (openPromptModalBtn && promptModal) {
    openPromptModalBtn.addEventListener('click', () => {
      promptModal.classList.add('open');
    });
  }

  if (closePromptModalBtn && promptModal) {
    closePromptModalBtn.addEventListener('click', () => {
      promptModal.classList.remove('open');
    });
  }

  if (promptModal) {
    promptModal.addEventListener('click', (e) => {
      if (e.target === promptModal) {
        promptModal.classList.remove('open');
      }
    });
  }

  // Student Prompt Modal controls
  const openStudentPromptBtn = document.getElementById('openStudentPromptBtn');
  const closeStudentPromptModalBtn = document.getElementById('closeStudentPromptModalBtn');
  const studentPromptModal = document.getElementById('studentPromptModal');

  if (openStudentPromptBtn && studentPromptModal) {
    openStudentPromptBtn.addEventListener('click', () => {
      studentPromptModal.classList.add('open');
    });
  }

  if (closeStudentPromptModalBtn && studentPromptModal) {
    closeStudentPromptModalBtn.addEventListener('click', () => {
      studentPromptModal.classList.remove('open');
    });
  }

  if (studentPromptModal) {
    studentPromptModal.addEventListener('click', (e) => {
      if (e.target === studentPromptModal) {
        studentPromptModal.classList.remove('open');
      }
    });
  }

  // Teacher Prompt Modal controls
  const openTeacherPromptBtn = document.getElementById('openTeacherPromptBtn');
  const closeTeacherPromptModalBtn = document.getElementById('closeTeacherPromptModalBtn');
  const teacherPromptModal = document.getElementById('teacherPromptModal');

  if (openTeacherPromptBtn && teacherPromptModal) {
    openTeacherPromptBtn.addEventListener('click', () => {
      teacherPromptModal.classList.add('open');
    });
  }

  if (closeTeacherPromptModalBtn && teacherPromptModal) {
    closeTeacherPromptModalBtn.addEventListener('click', () => {
      teacherPromptModal.classList.remove('open');
    });
  }

  if (teacherPromptModal) {
    teacherPromptModal.addEventListener('click', (e) => {
      if (e.target === teacherPromptModal) {
        teacherPromptModal.classList.remove('open');
      }
    });
  }

  // Employer Prompt Modal controls
  const openEmployerPromptBtn = document.getElementById('openEmployerPromptBtn');
  const closeEmployerPromptModalBtn = document.getElementById('closeEmployerPromptModalBtn');
  const employerPromptModal = document.getElementById('employerPromptModal');

  if (openEmployerPromptBtn && employerPromptModal) {
    openEmployerPromptBtn.addEventListener('click', () => {
      employerPromptModal.classList.add('open');
    });
  }

  if (closeEmployerPromptModalBtn && employerPromptModal) {
    closeEmployerPromptModalBtn.addEventListener('click', () => {
      employerPromptModal.classList.remove('open');
    });
  }

  if (employerPromptModal) {
    employerPromptModal.addEventListener('click', (e) => {
      if (e.target === employerPromptModal) {
        employerPromptModal.classList.remove('open');
      }
    });
  }

  // Business Prompt Modal controls
  const openBusinessPromptBtn = document.getElementById('openBusinessPromptBtn');
  const closeBusinessPromptModalBtn = document.getElementById('closeBusinessPromptModalBtn');
  const businessPromptModal = document.getElementById('businessPromptModal');

  if (openBusinessPromptBtn && businessPromptModal) {
    openBusinessPromptBtn.addEventListener('click', () => {
      businessPromptModal.classList.add('open');
    });
  }

  if (closeBusinessPromptModalBtn && businessPromptModal) {
    closeBusinessPromptModalBtn.addEventListener('click', () => {
      businessPromptModal.classList.remove('open');
    });
  }

  if (businessPromptModal) {
    businessPromptModal.addEventListener('click', (e) => {
      if (e.target === businessPromptModal) {
        businessPromptModal.classList.remove('open');
      }
    });
  }

  // Designer Prompt Modal controls
  const openDesignerPromptBtn = document.getElementById('openDesignerPromptBtn');
  const closeDesignerPromptModalBtn = document.getElementById('closeDesignerPromptModalBtn');
  const designerPromptModal = document.getElementById('designerPromptModal');

  if (openDesignerPromptBtn && designerPromptModal) {
    openDesignerPromptBtn.addEventListener('click', () => {
      designerPromptModal.classList.add('open');
    });
  }

  if (closeDesignerPromptModalBtn && designerPromptModal) {
    closeDesignerPromptModalBtn.addEventListener('click', () => {
      designerPromptModal.classList.remove('open');
    });
  }

  if (designerPromptModal) {
    designerPromptModal.addEventListener('click', (e) => {
      if (e.target === designerPromptModal) {
        designerPromptModal.classList.remove('open');
      }
    });
  }

  // Dynamic Demo Modal controls
  const clickableCards = document.querySelectorAll('.clickable-card');
  const demoPromptModal = document.getElementById('demoPromptModal');
  const demoModalTitle = document.getElementById('demoModalTitle');
  const demoModalPrompt = document.getElementById('demoModalPrompt');
  const demoModalDesc = document.getElementById('demoModalDesc');
  const closeDemoPromptModalBtn = document.getElementById('closeDemoPromptModalBtn');

  clickableCards.forEach(card => {
    card.addEventListener('click', () => {
      if (demoPromptModal) {
        demoModalTitle.textContent = card.dataset.demoTitle;
        demoModalPrompt.textContent = card.dataset.demoPrompt;
        demoModalDesc.innerHTML = card.dataset.demoDesc;
        demoPromptModal.classList.add('open');
      }
    });
  });

  if (closeDemoPromptModalBtn && demoPromptModal) {
    closeDemoPromptModalBtn.addEventListener('click', () => {
      demoPromptModal.classList.remove('open');
    });
  }

  if (demoPromptModal) {
    demoPromptModal.addEventListener('click', (e) => {
      if (e.target === demoPromptModal) {
        demoPromptModal.classList.remove('open');
      }
    });
  }

  if (demoModalPrompt) {
    demoModalPrompt.addEventListener('click', () => {
      navigator.clipboard.writeText(demoModalPrompt.textContent).then(() => {
        const originalTitle = demoModalPrompt.getAttribute('title') || 'Click to copy';
        demoModalPrompt.setAttribute('title', 'کۆپی کرا! ✓');
        // Simple visual feedback
        const originalBg = demoModalPrompt.style.background;
        demoModalPrompt.style.background = 'rgba(0, 242, 254, 0.2)';
        setTimeout(() => {
          demoModalPrompt.setAttribute('title', originalTitle);
          demoModalPrompt.style.background = originalBg;
        }, 1500);
      });
    });
  }

  // Copy button for Slide 16 modal
  const copyDemoPromptBtn = document.getElementById('copyDemoPromptBtn');
  if (copyDemoPromptBtn && demoModalPrompt) {
    copyDemoPromptBtn.addEventListener('click', () => {
      navigator.clipboard.writeText(demoModalPrompt.textContent).then(() => {
        const btnSpan = copyDemoPromptBtn.querySelector('span');
        const originalText = btnSpan.textContent;
        btnSpan.textContent = 'کۆپی کرا! ✓';
        copyDemoPromptBtn.style.background = 'rgba(16, 185, 129, 0.2)';
        copyDemoPromptBtn.style.borderColor = 'rgba(16, 185, 129, 0.4)';
        setTimeout(() => {
          btnSpan.textContent = originalText;
          copyDemoPromptBtn.style.background = '';
          copyDemoPromptBtn.style.borderColor = '';
        }, 1500);
      });
    });
  }

  // Initialize Copy Prompt Buttons for all level-boxes
  const levelBoxes = document.querySelectorAll('.level-box');
  levelBoxes.forEach(box => {
    if (box.id === 'demoModalPrompt') {
      return;
    }
    
    // Create copy button
    const btn = document.createElement('button');
    btn.className = 'copy-prompt-btn';
    btn.title = 'کۆپیکردنی پرۆمپت';
    btn.innerHTML = `
      <svg width="12" height="12" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
        <rect x="9" y="9" width="13" height="13" rx="2" ry="2"></rect>
        <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"></path>
      </svg>
      <span>کۆپی</span>
    `;
    
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      let textToCopy = "";
      box.childNodes.forEach(child => {
        if (child.nodeType === 3) { // Text Node
          textToCopy += child.nodeValue;
        } else if (child.nodeType === 1 && child.className !== 'copy-prompt-btn') {
          textToCopy += child.textContent;
        }
      });
      textToCopy = textToCopy.trim();
      
      navigator.clipboard.writeText(textToCopy).then(() => {
        btn.classList.add('copied');
        btn.querySelector('span').textContent = '✓ کۆپی کرا';
        setTimeout(() => {
          btn.classList.remove('copied');
          btn.querySelector('span').textContent = 'کۆپی';
        }, 1500);
      });
    });
    
    box.appendChild(btn);
  });

  // Handle Tab switches in both Modals individually
  const modalElements = document.querySelectorAll('.prompt-modal');
  modalElements.forEach(modal => {
    const tabBtns = modal.querySelectorAll('.tab-btn');
    const tabContents = modal.querySelectorAll('.tab-content');

    tabBtns.forEach(btn => {
      btn.addEventListener('click', () => {
        tabBtns.forEach(b => b.classList.remove('active'));
        tabContents.forEach(c => c.classList.remove('active'));

        btn.classList.add('active');
        const targetTabId = btn.dataset.tab || (btn.getAttribute('onclick') ? btn.getAttribute('onclick').match(/'([^']+)'/)?.[1] : null);
        if (targetTabId) {
          const targetContent = modal.querySelector(`#${targetTabId}`);
          if (targetContent) {
            targetContent.classList.add('active');
          }
        }
      });
    });
  });

  // Global switchTab helper
  window.switchTab = function(tabId) {
    const promptModal = document.getElementById('promptModal');
    if (!promptModal) return;
    const tabBtns = promptModal.querySelectorAll('.tab-btn');
    const tabContents = promptModal.querySelectorAll('.tab-content');

    tabBtns.forEach(b => {
      const bTab = b.dataset.tab || (b.getAttribute('onclick') ? b.getAttribute('onclick').match(/'([^']+)'/)?.[1] : null);
      if (bTab === tabId) {
        b.classList.add('active');
      } else {
        b.classList.remove('active');
      }
    });

    tabContents.forEach(c => {
      if (c.id === tabId) {
        c.classList.add('active');
      } else {
        c.classList.remove('active');
      }
    });
  };

  // Coffee Break 15-Minute Countdown Timer
  let coffeeTimerInterval = null;
  let coffeeSecondsLeft = 15 * 60; // 900 seconds

  window.toggleCoffeeTimer = function() {
    const display = document.getElementById('coffeeTimerDisplay');
    const btnSpan = document.querySelector('#startCoffeeTimerBtn span');

    if (coffeeTimerInterval) {
      clearInterval(coffeeTimerInterval);
      coffeeTimerInterval = null;
      if (btnSpan) btnSpan.textContent = 'بەردەوامبوون ▶';
    } else {
      if (coffeeSecondsLeft <= 0) coffeeSecondsLeft = 15 * 60;
      if (btnSpan) btnSpan.textContent = 'ڕاوەستان ⏸';

      coffeeTimerInterval = setInterval(() => {
        coffeeSecondsLeft--;
        if (coffeeSecondsLeft <= 0) {
          clearInterval(coffeeTimerInterval);
          coffeeTimerInterval = null;
          if (display) display.textContent = "00:00 - کاتی پشوو تەواو بوو! 🔔";
          if (btnSpan) btnSpan.textContent = 'دەستپێکردنەوە ▶';
          return;
        }
        const mins = Math.floor(coffeeSecondsLeft / 60).toString().padStart(2, '0');
        const secs = (coffeeSecondsLeft % 60).toString().padStart(2, '0');
        if (display) display.textContent = `${mins}:${secs}`;
      }, 1000);
    }
  };

  window.resetCoffeeTimer = function() {
    if (coffeeTimerInterval) {
      clearInterval(coffeeTimerInterval);
      coffeeTimerInterval = null;
    }
    coffeeSecondsLeft = 15 * 60;
    const display = document.getElementById('coffeeTimerDisplay');
    const btnSpan = document.querySelector('#startCoffeeTimerBtn span');
    if (display) display.textContent = "15:00";
    if (btnSpan) btnSpan.textContent = 'دەستپێکردنی کاتژمێر ▶';
  };

  // Initialize View
  updateSlideView();
});
