# Same as electron-builder's default finish page, except on an update (--updated):
# the branded progress page is shown, then the app relaunches by itself — no Finish click.
!macro customFinishPage
  !ifndef HIDE_RUN_AFTER_FINISH
    Function StartApp
      ${if} ${isUpdated}
        StrCpy $1 "--updated"
      ${else}
        StrCpy $1 ""
      ${endif}
      ${StdUtils.ExecShellAsUser} $0 "$launchLink" "open" "$1"
    FunctionEnd

    Function finishPreSkipIfUpdated
      ${if} ${isUpdated}
        Call StartApp
        Abort
      ${endif}
    FunctionEnd

    !define MUI_PAGE_CUSTOMFUNCTION_PRE finishPreSkipIfUpdated
    !define MUI_FINISHPAGE_RUN
    !define MUI_FINISHPAGE_RUN_FUNCTION "StartApp"
  !endif
  !insertmacro MUI_PAGE_FINISH
!macroend
