export type Designation = {
  designation: string;
  grade: string;
  cadre: string;
  approvedStrength: number;
  onRoll: number;
  status: "Active" | "Inactive";
};

export type Subcategory = {
  subcategory: string;
  designations: Designation[];
};

export type Category = {
  category: string;
  subcategories: Subcategory[];
};

export type StrengthData = {
  strengthStructure: Category[];
};

export const configurationNameKey = (name: string) => name.normalize("NFKC").trim().replace(/\s+/g, " ").toLowerCase();

export function validateStrengthData(data: StrengthData): void {
  if (!data || !Array.isArray(data.strengthStructure)) throw new Error("The configuration must contain a category list.");
  const validName = (value: unknown): value is string => typeof value === "string" && !!value.trim() && value.length <= 500;
  const categories = new Set<string>();
  for (const category of data.strengthStructure) {
    if (!category || !validName(category.category) || !Array.isArray(category.subcategories)) throw new Error("Every category needs a name and a subcategory list.");
    const key = configurationNameKey(category.category);
    if (categories.has(key)) throw new Error("This category already exists.");
    categories.add(key);
    const sections = new Set<string>();
    for (const section of category.subcategories) {
      if (!section || !validName(section.subcategory) || !Array.isArray(section.designations)) throw new Error("Every subcategory needs a name and a designation list.");
      const key = configurationNameKey(section.subcategory);
      if (sections.has(key)) throw new Error("This subcategory already exists in this category.");
      sections.add(key);
      for (const designation of section.designations) {
        if (!designation || ![designation.designation, designation.grade, designation.cadre].every(validName)) throw new Error("Designation, grade, and cadre are required.");
        if (![designation.approvedStrength, designation.onRoll].every((value) => Number.isSafeInteger(value) && value >= 0)) throw new Error("Strength values must be non-negative whole numbers.");
        if (designation.status !== "Active" && designation.status !== "Inactive") throw new Error("Choose Active or Inactive for designation status.");
      }
    }
  }
}

type DesignationInput = Omit<Designation, "status"> & { status?: Designation["status"] };

const emptyDesignation = (): Designation => ({
  designation: "",
  grade: "",
  cadre: "",
  approvedStrength: 0,
  onRoll: 0,
  status: "Active",
});

export const normalizeData = (data: StrengthData): StrengthData => ({
  strengthStructure: (data.strengthStructure ?? []).map((category) => ({
    category: category.category ?? "",
    subcategories: (category.subcategories ?? []).map((subcategory) => ({
      subcategory: subcategory.subcategory ?? "",
      designations: (subcategory.designations ?? []).map((designation) => ({
        designation: designation.designation ?? "",
        grade: designation.grade ?? "",
        cadre: designation.cadre ?? "",
        approvedStrength: Number(designation.approvedStrength) || 0,
        onRoll: Number(designation.onRoll) || 0,
        status: designation.status === "Inactive" ? "Inactive" : "Active",
      })),
    })),
  })),
});

export const addCategory = (data: StrengthData, category: string): StrengthData => {
  const name = category.trim();
  if (!name || data.strengthStructure.some((item) => item.category === name)) return data;
  return {
    ...data,
    strengthStructure: [...data.strengthStructure, { category: name, subcategories: [] }],
  };
};

export const addSubcategory = (
  data: StrengthData,
  categoryName: string,
  subcategory: string,
): StrengthData => {
  const name = subcategory.trim();
  if (!name) return data;
  return {
    ...data,
    strengthStructure: data.strengthStructure.map((category) => {
      if (category.category !== categoryName) return category;
      if (category.subcategories.some((item) => item.subcategory === name)) return category;
      return {
        ...category,
        subcategories: [...category.subcategories, { subcategory: name, designations: [] }],
      };
    }),
  };
};

export const addDesignation = (
  data: StrengthData,
  categoryName: string,
  subcategoryName: string,
  designation: DesignationInput,
): StrengthData => {
  const normalized = normalizeDesignationInput(designation);
  if (!normalized.designation.trim()) return data;
  return {
    ...data,
    strengthStructure: data.strengthStructure.map((category) => {
      if (category.category !== categoryName) return category;
      return {
        ...category,
        subcategories: category.subcategories.map((subcategory) => {
          if (subcategory.subcategory !== subcategoryName) return subcategory;
          return {
            ...subcategory,
            designations: [...subcategory.designations, normalized],
          };
        }),
      };
    }),
  };
};

export const updateDesignation = (
  data: StrengthData,
  categoryName: string,
  subcategoryName: string,
  designationName: string,
  next: DesignationInput,
): StrengthData => {
  const normalized = normalizeDesignationInput(next);
  if (!normalized.designation.trim()) return data;
  return {
    ...data,
    strengthStructure: data.strengthStructure.map((category) => {
      if (category.category !== categoryName) return category;
      return {
        ...category,
        subcategories: category.subcategories.map((subcategory) => {
          if (subcategory.subcategory !== subcategoryName) return subcategory;
          return {
            ...subcategory,
            designations: subcategory.designations.map((designation) =>
              designation.designation === designationName
                ? { ...normalized }
                : designation,
            ),
          };
        }),
      };
    }),
  };
};

export const deleteDesignation = (
  data: StrengthData,
  categoryName: string,
  subcategoryName: string,
  designationName: string,
): StrengthData => ({
  ...data,
  strengthStructure: data.strengthStructure.map((category) => {
    if (category.category !== categoryName) return category;
    return {
      ...category,
      subcategories: category.subcategories.map((subcategory) => {
        if (subcategory.subcategory !== subcategoryName) return subcategory;
        return {
          ...subcategory,
          designations: subcategory.designations.filter(
            (designation) => designation.designation !== designationName,
          ),
        };
      }),
    };
  }),
});

export const toggleDesignationStatus = (
  data: StrengthData,
  categoryName: string,
  subcategoryName: string,
  designationName: string,
): StrengthData => ({
  ...data,
  strengthStructure: data.strengthStructure.map((category) => {
    if (category.category !== categoryName) return category;
    return {
      ...category,
      subcategories: category.subcategories.map((subcategory) => {
        if (subcategory.subcategory !== subcategoryName) return subcategory;
        return {
          ...subcategory,
          designations: subcategory.designations.map((designation) =>
            designation.designation === designationName
              ? { ...designation, status: designation.status === "Active" ? "Inactive" : "Active" }
              : designation,
          ),
        };
      }),
    };
  }),
});

export const emptyDesignationForm = (): Designation => ({ ...emptyDesignation() });

export const normalizeDesignationInput = (input: DesignationInput): Designation => ({
  designation: input.designation.trim(),
  grade: input.grade.trim(),
  cadre: input.cadre.trim(),
  approvedStrength: Number(input.approvedStrength) || 0,
  onRoll: Number(input.onRoll) || 0,
  status: input.status === "Inactive" ? "Inactive" : "Active",
});
