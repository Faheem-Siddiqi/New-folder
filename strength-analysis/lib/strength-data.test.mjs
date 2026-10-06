import test from "node:test";
import assert from "node:assert/strict";

const { addCategory, addSubcategory, addDesignation, updateDesignation, toggleDesignationStatus } = await import("./strength-data.ts");

const makeTree = () => ({
  strengthStructure: [
    {
      category: "Production",
      subcategories: [
        {
          subcategory: "Ring",
          designations: [
            {
              designation: "Operator",
              grade: "E-03",
              cadre: "Worker",
              approvedStrength: 10,
              onRoll: 4,
              status: "Active",
            },
          ],
        },
      ],
    },
  ],
});

test("adds a category, subcategory, and designation", () => {
  const next = addDesignation(
    addSubcategory(addCategory(makeTree(), "Maintenance"), "Maintenance", "Spinning"),
    "Maintenance",
    "Spinning",
    {
      designation: "Fitter",
      grade: "E-03",
      cadre: "Worker",
      approvedStrength: 2,
      onRoll: 0,
      status: "Active",
    },
  );

  assert.equal(next.strengthStructure.at(-1)?.category, "Maintenance");
  assert.equal(next.strengthStructure.at(-1)?.subcategories.at(-1)?.subcategory, "Spinning");
  assert.equal(next.strengthStructure.at(-1)?.subcategories.at(-1)?.designations.at(-1)?.designation, "Fitter");
});

test("updates designation metadata and supports status toggling", () => {
  const initial = makeTree();
  const updated = updateDesignation(initial, "Production", "Ring", "Operator", {
    designation: "Senior Operator",
    grade: "E-04",
    cadre: "Staff",
    approvedStrength: 12,
    onRoll: 6,
    status: "Inactive",
  });

  assert.equal(updated.strengthStructure[0].subcategories[0].designations[0].designation, "Senior Operator");
  assert.equal(updated.strengthStructure[0].subcategories[0].designations[0].status, "Inactive");
  assert.equal(toggleDesignationStatus(updated, "Production", "Ring", "Senior Operator").strengthStructure[0].subcategories[0].designations[0].status, "Active");
});
