document.addEventListener('DOMContentLoaded', function () {
  const searchInput = document.querySelector('input.md-search__input');

  if (searchInput) {
    searchInput.addEventListener('keydown', function (e) {
      if (e.key === 'Enter') {
        const query = searchInput.value;

        fetch('https://script.google.com/macros/s/AKfycbxOXDOYrXaCQFCgnEsXW8OL7hZ0ZPLGU3pKY1mlHtLCF8JFmP014CtHFEhuRSWJJCr3/exec', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ query })
        });
      }
    });
  }
});
