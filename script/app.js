const scrollToTopBtn = document.getElementById("scrollToTop");

  // Показуємо кнопку, коли прокручування перевищує 300px
  window.addEventListener("scroll", () => {
    if (window.scrollY > 300) {
      scrollToTopBtn.style.display = "flex";
    } else {
      scrollToTopBtn.style.display = "none";
    }
  });

  // Додаємо плавний скрол до початку сторінки
  scrollToTopBtn.addEventListener("click", () => {
    window.scrollTo({
      top: 0,
      behavior: "smooth",
    });
  });

  // Отримуємо елементи
  const modal = document.getElementById("contactModal");
  const openButtons = document.querySelectorAll(".modal-btn"); // Кнопки "Зв'язатися"
  const closeModalBtn = document.querySelector(".close-btn");

  // Відкрити модалку при кліку на кнопку
  openButtons.forEach((button) => {
    button.addEventListener("click", () => {
      modal.style.display = "flex";
    });
  });

  // Закрити модалку при кліку на "x"
  closeModalBtn.addEventListener("click", () => {
    modal.style.display = "none";
  });

  // Закрити модалку при кліку за межі контенту
  window.addEventListener("click", (e) => {
    if (e.target === modal) {
      modal.style.display = "none";
    }
  });