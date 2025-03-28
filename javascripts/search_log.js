document.addEventListener('DOMContentLoaded', function () {
    const searchInput = document.querySelector('input.md-search__input');

    if (searchInput) {
        searchInput.addEventListener('keydown', function (e) {
            if (e.key === 'Enter') {
                const query = searchInput.value;

                fetch('https://script.google.com/macros/s/AKfycbzC7rgzGmi1u0-IzmpDDZin6iwocXaljPA8FpVH3vXwMmcXdDJ6DED--81lRyo5IV4-Pg/exec', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({query})
                })
                    .then(response => response.text())
                    .then(data => console.log("✅ 검색어 전송 성공:", data))
                    .catch(error => console.error("❌ 검색어 전송 실패:", error));
            }
        });
    }
});
