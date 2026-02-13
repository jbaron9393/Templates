const data = {
  "GI 1 and 2": [
    "GI 1 - Subcategory A",
    "GI 1 - Subcategory B",
    "GI 1 - Subcategory C",
    "GI 2 - Subcategory A",
    "GI 2 - Subcategory B",
    "GI 2 - Subcategory C"
  ]
};

const state = {
  selectedCategory: "",
  selectedSubcategory: ""
};

const categorySelect = document.getElementById("categorySelect");
const subcategoryContainer = document.getElementById("subcategoryContainer");
const output = document.getElementById("output");

function createChip(label, active, onClick) {
  const button = document.createElement("button");
  button.type = "button";
  button.className = `chip ${active ? "active" : ""}`;
  button.textContent = label;
  button.addEventListener("click", onClick);
  return button;
}

function renderCategoryOptions() {
  Object.keys(data).forEach((category) => {
    const option = document.createElement("option");
    option.value = category;
    option.textContent = category;
    categorySelect.append(option);
  });
}

function renderSubcategories() {
  subcategoryContainer.innerHTML = "";

  if (!state.selectedCategory) {
    subcategoryContainer.classList.add("empty-state");
    subcategoryContainer.textContent = "Select a category to load subcategories.";
    return;
  }

  const subcategories = data[state.selectedCategory] || [];
  if (!subcategories.length) {
    subcategoryContainer.classList.add("empty-state");
    subcategoryContainer.textContent = "No subcategories configured for this category.";
    return;
  }

  subcategoryContainer.classList.remove("empty-state");
  subcategories.forEach((subcategory) => {
    const chip = createChip(subcategory, state.selectedSubcategory === subcategory, () => {
      state.selectedSubcategory = subcategory;
      renderSubcategories();
      renderOutput();
    });
    subcategoryContainer.append(chip);
  });
}

function renderOutput() {
  if (!state.selectedCategory || !state.selectedSubcategory) {
    output.textContent = "Choose a category and subcategory.";
    return;
  }

  output.textContent = `${state.selectedCategory} > ${state.selectedSubcategory}`;
}

categorySelect.addEventListener("change", (event) => {
  state.selectedCategory = event.target.value;
  state.selectedSubcategory = "";
  renderSubcategories();
  renderOutput();
});

renderCategoryOptions();
renderSubcategories();
renderOutput();
