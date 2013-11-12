(function () {
var legacy_sections = {
};

    var h = document.location.hash;
    if (h && h.charAt(0) == '#' && legacy_sections.hasOwnProperty(h.slice(1)))
        document.location.hash = legacy_sections[h.slice(1)];
})();
