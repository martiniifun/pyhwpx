document.addEventListener('DOMContentLoaded', function () {
  const searchInput = document.querySelector('input.md-search__input');

  if (searchInput) {
    searchInput.addEventListener('keydown', function (e) {
      if (e.key === 'Enter') {
        const query = searchInput.value;

        fetch('https://script.google.com/macros/s/AKfycbzC7rgzGmi1u0-IzmpDDZin6iwocXaljPA8FpVH3vXwMmcXdDJ6DED--81lRyo5IV4-Pg/exec', {
          method: 'POST',
          mode: 'no-cors',
          headers: { 'Content-Type': 'text/plain' },  // 더 안전하게
          body: query  // JSON.stringify({ query }) 대신 문자열만
        });

        console.log("📤 검색어 전송 시도:", query);
      }
    });
  }
});
