(function () {
  "use strict";

  var app = document.getElementById("app");
  var config = window.BIRTHDAY_CLASS_CONFIG;

  function escapeHtml(value) {
    return String(value).replace(/[&<>"']/g, function (character) {
      return {
        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;",
        '"': "&quot;",
        "'": "&#39;"
      }[character];
    });
  }

  function portrait(classItem, extraClass) {
    return '<div class="portrait portrait--' + escapeHtml(classItem.portrait) + ' ' + extraClass + '" aria-hidden="true">' +
      '<span class="portrait__light"></span><span class="portrait__body"></span><span class="portrait__neck"></span><span class="portrait__head"></span><span class="portrait__hair"></span><span class="portrait__grain"></span></div>';
  }

  function renderHome() {
    var cards = config.classes.map(function (classItem, index) {
      return '<a class="class-card" href="#class/' + encodeURIComponent(classItem.id) + '" style="--card-order:' + index + '">' +
        portrait(classItem, "class-card__portrait") +
        '<div class="class-card__copy"><span class="eyebrow">' + escapeHtml(classItem.role) + '</span>' +
        '<h3>' + escapeHtml(classItem.title) + '</h3><span class="class-card__instructor">With ' + escapeHtml(classItem.instructor) + '</span>' +
        '<span class="class-card__arrow" aria-hidden="true">↗</span></div></a>';
    }).join("");

    app.innerHTML = '<section class="hero"><div class="hero__copy"><p class="eyebrow hero__eyebrow">A VERY SPECIAL EDITION · EST. YOUR BIRTHDAY</p>' +
      '<h1>Learn from<br>the people<br>who <em>love you.</em></h1>' +
      '<p class="hero__tagline">A small collection of big feelings, questionable expertise, and classes made just for ' + escapeHtml(config.recipientName) + '.</p>' +
      '<a class="button-link" href="#classes">EXPLORE THE CLASSES <span aria-hidden="true">↓</span></a>' +
      '<p class="hero__note">No homework. Snacks strongly encouraged.</p></div>' +
      '<div class="hero__art" aria-hidden="true"><div class="hero__frame"><div class="hero__sun"></div><div class="hero__shadow"></div><div class="hero__bloom"></div><span class="hero__art-caption">A GATHERING OF GREAT MINDS<br>(AND ONE HOUSEPLANT)</span><span class="hero__index">01 / 06</span></div></div>' +
      '<div class="hero__bottom"><span>AN ORIGINAL BIRTHDAY PRODUCTION</span><span>SCROLL TO FEEL APPRECIATED <b>↓</b></span></div></section>' +
      '<section class="catalog" id="classes"><div class="section-heading"><div><p class="eyebrow">YOUR FACULTY AWAITS</p><h2>Six classes.<br><em>One very lucky student.</em></h2></div><p class="section-heading__aside">A carefully assembled syllabus from people who know you best. Academic credentials: enthusiastic.</p></div>' +
      '<div class="class-grid">' + cards + '</div></section>' +
      '<section class="closing-note"><span class="closing-note__mark" aria-hidden="true">B.</span><p>Consider this your official permission to be celebrated.</p><span>WITH LOVE, YOUR VERY UNQUALIFIED FACULTY</span></section>';

    app.focus({ preventScroll: true });
  }

  function renderVideo(classItem) {
    if (!classItem.videoUrl) {
      return '<div class="video-placeholder"><span class="video-placeholder__mark" aria-hidden="true">B.</span><p class="eyebrow">THE STUDIO IS PREPARING</p><h2>Class coming soon</h2><p>Our instructor is rehearsing, locating the good light, or making everyone a snack.</p><span class="video-placeholder__bottom">PLEASE ENJOY THIS MOMENT OF ANTICIPATION</span></div>';
    }

    return '<video class="class-video" controls playsinline preload="metadata" aria-label="' + escapeHtml(classItem.title) + '"><source src="' + escapeHtml(classItem.videoUrl) + '">Your browser does not support video playback.</video>';
  }

  function renderClass(classItem) {
    var lessons = classItem.lessons.map(function (lesson, index) {
      return '<li><span class="lesson-number">0' + (index + 1) + '</span><span>' + escapeHtml(lesson) + '</span><span class="lesson-duration">A MOMENT WELL SPENT</span></li>';
    }).join("");

    app.innerHTML = '<section class="class-detail"><a class="back-link" href="#classes"><span aria-hidden="true">←</span> ALL CLASSES</a>' +
      '<div class="class-detail__intro"><div class="class-detail__title"><p class="eyebrow">A BIRTHDAYCLASS ORIGINAL</p><h1>' + escapeHtml(classItem.title) + '</h1><p class="class-detail__byline">A class with <strong>' + escapeHtml(classItem.instructor) + '</strong></p></div>' +
      portrait(classItem, "class-detail__portrait") + '</div>' +
      '<div class="class-detail__body"><div class="class-detail__main"><div class="video-frame">' + renderVideo(classItem) + '</div>' +
      '<div class="class-description"><p class="eyebrow">ABOUT THIS CLASS</p><p>' + escapeHtml(classItem.description) + '</p></div>' +
      '<section class="lesson-section"><p class="eyebrow">THE CURRICULUM</p><h2>In this class</h2><ol class="lesson-list">' + lessons + '</ol></section></div>' +
      '<aside class="instructor-bio"><p class="eyebrow">MEET YOUR INSTRUCTOR</p><h2>' + escapeHtml(classItem.instructor) + '</h2><p class="instructor-bio__role">' + escapeHtml(classItem.role) + '</p><p>' + escapeHtml(classItem.biography) + '</p><a class="text-link" href="#classes">MEET THE REST OF THE FACULTY <span aria-hidden="true">→</span></a></aside></div>' +
      '<div class="class-detail__signoff"><span>MADE ESPECIALLY FOR ' + escapeHtml(config.recipientName).toUpperCase() + '</span><span>WITH A QUESTIONABLE AMOUNT OF LOVE</span></div></section>';

    app.focus({ preventScroll: true });
  }

  function renderRoute() {
    var route = window.location.hash.replace(/^#/, "");
    var match = route.match(/^class\/([^/]+)$/);
    var classItem = match && config.classes.find(function (item) {
      return item.id === decodeURIComponent(match[1]);
    });

    if (classItem) {
      renderClass(classItem);
    } else {
      renderHome();
      if (route === "classes") {
        document.getElementById("classes").scrollIntoView();
      }
    }
  }

  window.addEventListener("hashchange", renderRoute);
  renderRoute();
})();
