# AGENTS.md

IMPORTANT: Prefer retrieval-led reasoning over pre-training-led reasoning for any React, Next.js, or UI tasks. Consult the skill docs below before relying on training data.

## Installed Skills Index

root: ./.agents/skills

| Skill | Description | Entry Point | Rules |
|---|---|---|---|
| vercel-react-best-practices | React/Next.js performance optimization (57 rules, 8 categories). Use when writing, reviewing, or refactoring React/Next.js code. | SKILL.md, AGENTS.md | rules/{async-,bundle-,server-,client-,rerender-,rendering-,js-,advanced-}*.md |
| vercel-composition-patterns | React composition patterns (8 rules, 4 categories). Use when refactoring boolean prop proliferation, building compound components, or designing reusable APIs. Includes React 19 changes. | SKILL.md, AGENTS.md | rules/{architecture-,state-,patterns-,react19-}*.md |
| web-design-guidelines | Web Interface Guidelines compliance review. Use when asked to review UI, check accessibility, audit design, or review UX. Fetches latest rules from remote source. | SKILL.md | (fetched at runtime from https://raw.githubusercontent.com/vercel-labs/web-interface-guidelines/main/command.md) |

## Skills File Map

```
.agents/skills/
|vercel-react-best-practices:{SKILL.md,AGENTS.md}
|vercel-react-best-practices/rules:{async-defer-await.md,async-parallel.md,async-dependencies.md,async-api-routes.md,async-suspense-boundaries.md,bundle-barrel-imports.md,bundle-dynamic-imports.md,bundle-defer-third-party.md,bundle-conditional.md,bundle-preload.md,server-auth-actions.md,server-cache-react.md,server-cache-lru.md,server-dedup-props.md,server-serialization.md,server-parallel-fetching.md,server-after-nonblocking.md,client-swr-dedup.md,client-event-listeners.md,client-passive-event-listeners.md,client-localstorage-schema.md,rerender-defer-reads.md,rerender-memo.md,rerender-memo-with-default-value.md,rerender-dependencies.md,rerender-derived-state.md,rerender-derived-state-no-effect.md,rerender-functional-setstate.md,rerender-lazy-state-init.md,rerender-simple-expression-in-memo.md,rerender-move-effect-to-event.md,rerender-transitions.md,rerender-use-ref-transient-values.md,rendering-animate-svg-wrapper.md,rendering-content-visibility.md,rendering-hoist-jsx.md,rendering-svg-precision.md,rendering-hydration-no-flicker.md,rendering-hydration-suppress-warning.md,rendering-activity.md,rendering-conditional-render.md,rendering-usetransition-loading.md,js-batch-dom-css.md,js-index-maps.md,js-cache-property-access.md,js-cache-function-results.md,js-cache-storage.md,js-combine-iterations.md,js-length-check-first.md,js-early-exit.md,js-hoist-regexp.md,js-min-max-loop.md,js-set-map-lookups.md,js-tosorted-immutable.md,advanced-event-handler-refs.md,advanced-init-once.md,advanced-use-latest.md}
|vercel-composition-patterns:{SKILL.md,AGENTS.md}
|vercel-composition-patterns/rules:{architecture-avoid-boolean-props.md,architecture-compound-components.md,state-decouple-implementation.md,state-context-interface.md,state-lift-state.md,patterns-explicit-variants.md,patterns-children-over-render-props.md,react19-no-forwardref.md}
|web-design-guidelines:{SKILL.md}
```

## Priority Quick Reference

When working on React/Next.js code, prioritize by impact:

1. **CRITICAL** -- Eliminate waterfalls (async-*), reduce bundle size (bundle-*)
2. **HIGH** -- Server-side performance (server-*), component architecture (architecture-*)
3. **MEDIUM** -- Client data fetching (client-*), re-renders (rerender-*), rendering (rendering-*), state management (state-*), composition patterns (patterns-*, react19-*)
4. **LOW** -- JS micro-optimizations (js-*), advanced patterns (advanced-*)

Read the specific rule file for detailed explanations and correct/incorrect code examples before generating or refactoring code.
