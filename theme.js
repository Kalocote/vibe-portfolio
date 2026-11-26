/* 深浅色切换：选择记在本地，所有页面统一生效 */
(function () {
    const KEY = "vibe-theme";
    const btn = document.querySelector("#themeBtn");

    function apply(dark) {
        document.body.classList.toggle("dark", dark);
        if (btn) btn.textContent = dark ? "浅色" : "深色";
    }

    // 先读取上次保存的选择，没有就默认浅色
    apply(localStorage.getItem(KEY) === "dark");

    if (btn) {
        btn.addEventListener("click", function () {
            const dark = !document.body.classList.contains("dark");
            localStorage.setItem(KEY, dark ? "dark" : "light");
            apply(dark);
        });
    }
})();
