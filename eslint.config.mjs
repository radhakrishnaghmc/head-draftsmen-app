import js from '@eslint/js'
import tseslint from 'typescript-eslint'
import reactHooks from 'eslint-plugin-react-hooks'
import prettier from 'eslint-config-prettier'
import globals from 'globals'

export default tseslint.config(
  { ignores: ['node_modules', 'out', 'dist', 'release', 'resources', 'graphify-out', 'src/_tmp', 'patches', 'scripts'] },
  js.configs.recommended,
  ...tseslint.configs.recommended,
  {
    files: ['src/**/*.{ts,tsx}'],
    plugins: { 'react-hooks': reactHooks },
    languageOptions: { globals: globals.browser },
    rules: {
      ...reactHooks.configs.recommended.rules,
      // Stricter React-compiler rules, as warnings so existing code passes today.
      'react-hooks/set-state-in-effect': 'warn',
      'react-hooks/refs': 'warn',
      'react-hooks/purity': 'warn'
    }
  },
  {
    files: ['electron/**/*.ts', 'core/**/*.ts', 'tests/**/*.ts'],
    languageOptions: { globals: { ...globals.node } }
  },
  {
    linterOptions: { reportUnusedDisableDirectives: 'off' },
    rules: {
      // TypeScript already reports undefined names; the core rule only gives false positives on types/globals.
      'no-undef': 'off',
      // Stricter rules, surfaced as warnings so the existing code base passes today and they can be fixed gradually.
      'no-useless-escape': 'warn',
      'no-useless-assignment': 'warn',
      'no-control-regex': 'warn',
      'preserve-caught-error': 'warn',
      '@typescript-eslint/no-require-imports': 'warn',
      '@typescript-eslint/no-unused-expressions': 'warn',
      // Existing code uses a handful of deliberate `any`s and `_`-prefixed unused names.
      '@typescript-eslint/no-explicit-any': 'warn',
      '@typescript-eslint/no-unused-vars': ['warn', { argsIgnorePattern: '^_', varsIgnorePattern: '^_' }]
    }
  },
  prettier
)
