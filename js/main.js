document.addEventListener('DOMContentLoaded', function () {

  // ========== 导航栏滚动效果 ==========
  var header = document.querySelector('.site-header');
  window.addEventListener('scroll', function () {
    if (window.scrollY > 20) {
      header.classList.add('scrolled');
    } else {
      header.classList.remove('scrolled');
    }
  });

  // ========== 移动端菜单切换 ==========
  var menuToggle = document.querySelector('.menu-toggle');
  var mainNav = document.querySelector('.main-nav');

  if (menuToggle && mainNav) {
    menuToggle.addEventListener('click', function () {
      var isOpen = mainNav.classList.toggle('open');
      menuToggle.setAttribute('aria-expanded', isOpen);
    });

    mainNav.querySelectorAll('a').forEach(function (link) {
      link.addEventListener('click', function () {
        mainNav.classList.remove('open');
        menuToggle.setAttribute('aria-expanded', 'false');
      });
    });
  }

  // ========== FAQ 手风琴 ==========
  document.querySelectorAll('.faq-question').forEach(function (btn) {
    btn.addEventListener('click', function () {
      var item = btn.closest('.faq-item');
      var answer = item.querySelector('.faq-answer');
      var isActive = item.classList.contains('active');

      document.querySelectorAll('.faq-item.active').forEach(function (openItem) {
        openItem.classList.remove('active');
        openItem.querySelector('.faq-answer').style.maxHeight = null;
        openItem.querySelector('.faq-question').setAttribute('aria-expanded', 'false');
      });

      if (!isActive) {
        item.classList.add('active');
        answer.style.maxHeight = answer.scrollHeight + 'px';
        btn.setAttribute('aria-expanded', 'true');
      }
    });
  });

  // ========== 滚动渐入动画 ==========
  var fadeElements = document.querySelectorAll('.fade-in');

  if ('IntersectionObserver' in window) {
    var observer = new IntersectionObserver(function (entries) {
      entries.forEach(function (entry) {
        if (entry.isIntersecting) {
          entry.target.classList.add('visible');
          observer.unobserve(entry.target);
        }
      });
    }, {
      threshold: 0.1,
      rootMargin: '0px 0px -40px 0px'
    });

    fadeElements.forEach(function (el) {
      observer.observe(el);
    });
  } else {
    fadeElements.forEach(function (el) {
      el.classList.add('visible');
    });
  }

  // ========== 联系表单处理 ==========
  var contactForm = document.getElementById('contactForm');
  if (contactForm) {
    contactForm.addEventListener('submit', function (e) {
      e.preventDefault();

      var name = document.getElementById('name');
      var phone = document.getElementById('phone');
      var message = document.getElementById('message');
      var isValid = true;

      [name, phone, message].forEach(function (field) {
        field.style.borderColor = '';
      });

      if (!name.value.trim()) {
        name.style.borderColor = '#ef4444';
        isValid = false;
      }
      if (!phone.value.trim() || !/^1[3-9]\d{9}$/.test(phone.value.trim())) {
        phone.style.borderColor = '#ef4444';
        isValid = false;
      }
      if (!message.value.trim()) {
        message.style.borderColor = '#ef4444';
        isValid = false;
      }

      if (isValid) {
        showToast('感谢您的咨询！我们的顾问将在24小时内与您联系。');
        contactForm.reset();
      } else {
        showToast('请完整填写必填项（姓名、手机号、需求描述）');
      }
    });
  }

  // ========== Toast 提示 ==========
  function showToast(msg) {
    var existing = document.querySelector('.toast');
    if (existing) existing.remove();

    var toast = document.createElement('div');
    toast.className = 'toast';
    toast.textContent = msg;
    toast.setAttribute('role', 'alert');
    document.body.appendChild(toast);

    requestAnimationFrame(function () {
      toast.classList.add('show');
    });

    setTimeout(function () {
      toast.classList.remove('show');
      setTimeout(function () {
        toast.remove();
      }, 400);
    }, 4000);
  }

  // ========== 平滑滚动到锚点 ==========
  document.querySelectorAll('a[href^="#"]').forEach(function (anchor) {
    anchor.addEventListener('click', function (e) {
      var targetId = this.getAttribute('href');
      if (targetId === '#') return;

      var target = document.querySelector(targetId);
      if (target) {
        e.preventDefault();
        var headerHeight = header.offsetHeight;
        var targetPos = target.getBoundingClientRect().top + window.pageYOffset - headerHeight - 20;
        window.scrollTo({ top: targetPos, behavior: 'smooth' });
      }
    });
  });

});
