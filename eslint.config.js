const spfxProfile = require('@microsoft/eslint-config-spfx/lib/flat-profiles/react');

module.exports = [
  ...spfxProfile,
  {
    files: ['**/*.ts', '**/*.tsx'],
    languageOptions: {
      parserOptions: {
        tsconfigRootDir: __dirname,
        project: './tsconfig.json'
      }
    },
    rules: {
      // null is used throughout for "no selection" UI state — safe in React 17
      '@rushstack/no-new-null': 'off',
      // OData field names must use bracket notation; dot-notation false positive
      'dot-notation': 'off',
      // void is used for intentional fire-and-forget async calls
      'no-void': 'off',
      // JSX string literals with quotes are acceptable
      'react/no-unescaped-entities': 'off',
      // ReactDOM.render/unmount pairing is handled by onDispose
      '@rushstack/pair-react-dom-render-unmount': 'off',
      // Explicit return types on every inline handler is overly verbose
      '@typescript-eslint/explicit-function-return-type': 'off',
      // != null intentionally catches both null and undefined in OData checks
      'eqeqeq': 'off',
      // Constants forward-referenced across hoisted function bodies are safe
      '@typescript-eslint/no-use-before-define': 'off',
      // Unused var detection is handled by TypeScript's own noUnusedLocals
      '@typescript-eslint/no-unused-vars': 'off'
    }
  }
];
