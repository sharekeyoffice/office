// wrapper-customization.js — editorConfig builder for v1 wrapper
//
// Exports buildEditorConfig(type) on window. `type` is 'word' | 'cell' | 'slide'.
// Used by wrapper-mount.js (parent context) when responding to onAppReady.
//
// Decisions encoded here:
//   - Plugins subsystem disabled (kills marketplace + AI plugins). 10.0.2 scope-decision.
//   - Protect tab disabled (signing/passwords irrelevant under E2E main app). 10.0.2.
//   - Comments disabled (we'll build main app-integrated commenting later).
//   - Mail merge / forms creator hidden (out of v1 scope).
//   - English locale only.
//   - canCoAuthoring=false: single-editor mode for v1; v2 adds collab.
//   - Mode is ALWAYS 'edit' here. Read-only state is applied dynamically via
//     `asc_setRestriction(View)` from wrapper-mount.js after onAppReady. This
//     lets the main app hot-toggle view↔edit without destroying the editor.
//   - Autosave is forced OFF because we drive our own autosave via
//     wrapper-postmessage.js; OnlyOffice's built-in autosave would try to
//     talk to DocServer (which we don't run).

(function (global) {
  'use strict';

  var DOCUMENT_TYPE = { word: 'word', cell: 'cell', slide: 'slide' };

  var FALLBACK_USER = { id: 'sk-editor-user', name: 'Sharekey user' };

  function buildEditorConfig(type, opts) {
    if (!DOCUMENT_TYPE[type]) {
      throw new Error('wrapper-customization: unknown document type: ' + type);
    }
    opts = opts || {};

    return {
      // Document descriptor — populated for real on document load (10.1.2);
      // for 10.1.1's UI smoke test we ship a minimal placeholder.
      document: {
        title:    opts.title    || 'Untitled',
        url:      opts.url      || '',
        fileType: opts.fileType || ({ word: 'docx', cell: 'xlsx', slide: 'pptx' }[type]),
        key:      opts.key      || ('wrap-' + Date.now()),
        permissions: {
          edit:                   true,    // always grant edit; view-only is applied via asc_setRestriction
          download:               true,
          review:                 false,   // collab feature, not v1
          print:                  false,   // out of v1 scope
          comment:                false,   // disabled in customization too
          modifyContentControl:   false,
          modifyFilter:           true,    // allow column auto-filter/sort (cell editor); Main.js: canModifyFilter = (modifyFilter!==false)
          fillForms:              false,
          copy:                   true,    // both modes can copy
          chat:                   false,
          protect:                false    // hard-off, see customization.layout.toolbar.protect
        }
      },

      editorConfig: {
        mode:                 'edit',       // always boot in edit mode; restriction applied dynamically
        lang:                 'en-US',
        canCoAuthoring:       false,        // v1 = single-editor; v2 = collab
        canBackToFolder:      false,
        canRequestClose:      true,         // host can ask us to close

        user: {
          id:    opts.userId   || FALLBACK_USER.id,
          name:  opts.userName || FALLBACK_USER.name,
          group: ''
        },

        customization: {
          // ---- Branding (ONLYOFFICE trademark policy — brand MUST read
          //      "Sharekey", never "ONLYOFFICE"). Header.setBranding() swaps the
          //      #header-logo for this image when asc_getCustomization() is true
          //      (shimmed in wrapper-boot.js). `image` is used for both header
          //      themes via getSuitableLogo's fallback; add imageLight if a
          //      light-on-dark variant is needed for a dark header.
          //      url:'' → logo is not a clickable link. Asset served from
          //      overlay/icons/ at /icons/sharekey-logo.svg.
          logo: {
            image:     '/icons/sharekey-logo.svg',
            imageDark: '/icons/sharekey-logo.svg',
            url:       ''
          },

          // ---- license-INDEPENDENT knobs (parsed in Main.js) -----------------
          plugins:        false,    // kills marketplace + AI plugins (incl. Plugins tab)
          macros:         false,    // kills macros + MacrosAiDialog
          comments:       false,    // kills comments subsystem (LeftMenu + insert + panel)
          mentionShare:   false,
          compactToolbar: false,
          toolbarNoTabs:  false,
          anonymous:      { request: false },
          features: {
            roles:                   false,
            tabBackground:           'header',
            // Suppress the "New <feature> … Got it" onboarding tooltips
            // (the synch-tip-root popups, e.g. "Multipage view"). Each editor's
            // Toolbar.setMode registers its isNewFeature tips only when
            // FeaturesManager.isFeatureEnabled('featuresTips', true) is truthy
            // (→ features.featuresTips !== false). Setting it false skips that
            // addTips() call entirely, so the intro tips are never registered.
            // The functional warnings (disconnect / refreshFile / sessionIdle /
            // cantModifyFilter / copyDisabled / licenseError) live in a separate,
            // un-gated addTips block and are unaffected.
            featuresTips:            false
          },
          help:           false,    // hide the in-app help link (we trimmed help docs anyway)
          about:          false,
          autosave:       false,    // we drive autosave via wrapper-postmessage.js; SDK autosave talks to DocServer
          forcesave:      false,
          chat:           false,

          // ---- license-GATED knobs (require asc_getCanBranding shim) ---------
          // The shim in wrapper-boot.js makes asc_getCanBranding return true;
          // the LayoutManager then walks this tree and hides matching elements.
          layout: {
            toolbar: {
              file: { close: false },
              protect: false,           // out of v1 scope
              draw:    false,           // not needed for main app embed
              plugins: false,           // matches plugins:false above; redundant safety
              collaboration: false      // hide the whole Collaboration/Review tab
                                        // (layoutname 'toolbar-collaboration'); the
                                        // tab is otherwise shown because the doc is
                                        // editable. `false` (not {mailmerge:false})
                                        // makes LayoutManager.isElementVisible(
                                        // 'toolbar-collaboration') return false.
            },
            header: {
              user:  false,            // hide the current-user "E" avatar
                                        // (slot data-layout-name="header-user",
                                        // the .color-user-name initials button;
                                        // Header.js setUserName). Identity is
                                        // implicit in the main app embed.
              users: false             // hide the co-authoring users dropdown
                                        // (data-layout-name="header-users", the
                                        // svg-icon-users button) — no co-authors
                                        // in single-editor v1.
                                        //
                                        // LayoutManager.applyCustomization()
                                        // turns each header:{<k>:false} into the
                                        // exact selector [data-layout-name=header-<k>]
                                        // and .hide()s it. "header-user" and
                                        // "header-users" are distinct exact matches.
            },
            leftMenu: {
              comments:  false,
              chat:      false,
              navigation: true,
              support:   false,
              about:     false,
              feedback:  false
            },
            rightMenu: true,
            statusBar: {
              docLang:  false,
              textLang: false,
              actionStatus: false
            }
          }
        }
      }
    };
  }

  global.buildEditorConfig = buildEditorConfig;
})(window);
