Office.onReady(() => {
  const btn = document.getElementById("btnRefresh");
  if (btn) btn.addEventListener("click", () => location.reload());
});
