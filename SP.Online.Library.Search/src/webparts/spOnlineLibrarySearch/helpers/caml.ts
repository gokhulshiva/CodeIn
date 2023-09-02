export enum MergeType { Or, And }

export function MergeCAMLConditions(conditions: any[], type: MergeType): string {
    if (conditions.length == 0) return "";

    const typeStart = (type == MergeType.And ? "<And>" : "<Or>");
    const typeEnd = (type == MergeType.And ? "</And>" : "</Or>");

    // Build hierarchical structure
    while (conditions.length >= 2) {
        const complexConditions = [];

        for (let i = 0; i < conditions.length; i += 2) {
            if (conditions.length == i + 1) // Only one condition left
                complexConditions.push(conditions[i]);
            else // Two condotions - merge
                complexConditions.push(typeStart + conditions[i] + conditions[i + 1] + typeEnd);
        }

        conditions = complexConditions;
    }

    return conditions[0];
}