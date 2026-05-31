(function () {
  function initializeReveal() {
    if (!window.Reveal) return;

    Reveal.initialize({
      hash: true,
      controls: true,
      progress: true,
      center: false,
      width: 1280,
      height: 720,
      margin: 0.03,
      transition: "slide",
      backgroundTransition: "fade",
      plugins: window.RevealNotes ? [RevealNotes] : []
    });
  }

  function initializeIcons() {
    if (window.lucide) {
      window.lucide.createIcons({
        attrs: {
          "stroke-width": 2.2
        }
      });
    }
  }

  function initializeXmlWriteReveal() {
    const demo = document.querySelector("[data-xmlwrite-demo]");
    const button = document.querySelector("[data-xmlwrite-button]");
    if (!demo || !button) return;

    button.addEventListener("click", () => {
      const writing = demo.classList.toggle("writing");
      button.querySelector("span").textContent = writing ? "처음으로" : "이 문장을 입력하면?";
    });
  }

  function initializeZipReveal() {
    const demo = document.querySelector("[data-zip-demo]");
    const button = document.querySelector("[data-zip-button]");
    if (!demo || !button) return;

    const label = button.querySelector("span");
    button.addEventListener("click", () => {
      if (!demo.classList.contains("renamed")) {
        demo.classList.add("renamed");
        label.textContent = "② 압축 풀어보기";
      } else if (!demo.classList.contains("opened")) {
        demo.classList.add("opened");
        label.textContent = "처음부터 다시";
      } else {
        demo.classList.remove("renamed", "opened");
        label.textContent = "① 확장자를 .zip으로 바꾸기";
      }
    });
  }

  function initializeMediaReveal() {
    const demo = document.querySelector("[data-media-demo]");
    const mediaRow = document.querySelector("[data-open-media]");
    const back = document.querySelector("[data-media-back]");
    if (!demo || !mediaRow) return;

    mediaRow.addEventListener("click", () => demo.classList.add("open"));
    if (back) back.addEventListener("click", () => demo.classList.remove("open"));
  }

  function initializeProtectReveal() {
    const demo = document.querySelector("[data-protect-demo]");
    const button = document.querySelector("[data-protect-button]");
    if (!demo || !button) return;

    button.addEventListener("click", () => {
      const unlocked = demo.classList.toggle("unlocked");
      button.querySelector("span").textContent = unlocked ? "보호 태그 되돌리기" : "보호 태그 지우기";
    });
  }

  function initializeBatchReveal() {
    const demo = document.querySelector("[data-batch-demo]");
    const button = document.querySelector("[data-batch-button]");
    if (!demo || !button) return;

    button.addEventListener("click", () => {
      const stuck = demo.classList.toggle("stuck");
      button.querySelector("span").textContent = stuck ? "처음으로" : "10개 파일 한 번에 처리";
    });
  }

  document.addEventListener("DOMContentLoaded", () => {
    initializeReveal();
    initializeIcons();
    initializeXmlWriteReveal();
    initializeZipReveal();
    initializeMediaReveal();
    initializeProtectReveal();
    initializeBatchReveal();
  });
})();
