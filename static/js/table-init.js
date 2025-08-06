document.addEventListener("DOMContentLoaded", () => {
  $("table").DataTable({
    paging: true,
    searching: true,
    scrollX: true,
    fixedHeader: true
  });
});
