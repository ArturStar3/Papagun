import type { ArticleGroup, Catalog } from './selection';

export type SelectionTemplate = {
  id: string;
  name: string;
  months: string[];
  branches: string[];
  articles: string[];
  groups: ArticleGroup[];
};

const storageKey = 'papagun-selection-templates-v1';

function strings(value: unknown): value is string[] {
  return Array.isArray(value) && value.every(item => typeof item === 'string');
}

function validTemplate(value: unknown): value is SelectionTemplate {
  if (!value || typeof value !== 'object') return false;
  const item = value as Partial<SelectionTemplate>;
  return typeof item.id === 'string' && typeof item.name === 'string' &&
    strings(item.months) && strings(item.branches) && strings(item.articles) &&
    Array.isArray(item.groups) && item.groups.every(group => group &&
      typeof group.id === 'string' && typeof group.title === 'string' && strings(group.articles));
}

export function readTemplates(): SelectionTemplate[] {
  try {
    const value: unknown = JSON.parse(localStorage.getItem(storageKey) || '[]');
    return Array.isArray(value) ? value.filter(validTemplate) : [];
  } catch {
    return [];
  }
}

export function writeTemplates(items: SelectionTemplate[]): void {
  localStorage.setItem(storageKey, JSON.stringify(items));
}

export function templateProblem(template: SelectionTemplate, catalog: Catalog): string | null {
  const months = new Set(catalog.availableMonths);
  const branches = new Set(catalog.availableBranches);
  const articles = new Map(catalog.articles.map(item => [item.id, item]));
  if (!template.months.length || template.months.some(item => !months.has(item))) return 'В шаблоне есть недоступные месяцы.';
  if (!template.branches.length || template.branches.some(item => !branches.has(item))) return 'В шаблоне есть недоступные филиалы.';
  if (new Set(template.months).size !== template.months.length ||
      new Set(template.branches).size !== template.branches.length) return 'В шаблоне повторяется месяц или филиал.';
  const used = [...template.articles, ...template.groups.flatMap(group => group.articles)];
  if (used.some(id => !articles.has(id))) return 'В шаблоне есть статьи, отсутствующие в текущих книгах.';
  if (used.some(id => ['211', '212', '213'].includes(articles.get(id)?.code || ''))) return 'В шаблоне выбрана итоговая строка раздела.';
  if (new Set(used).size !== used.length) return 'В шаблоне повторяется статья.';
  if (template.groups.some(group => !group.title.trim() || group.articles.length < 2 ||
      new Set(group.articles.map(id => articles.get(id)?.group)).size !== 1)) {
    return 'В шаблоне есть некорректная группа статей.';
  }
  return null;
}
