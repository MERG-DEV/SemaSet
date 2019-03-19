/***************************************************************
 * Name:      MainFrame.cpp
 * Purpose:   Code for Application Frame
 * Author:    Chris White (whitecf@bcs.org.uk)
 * Created:   2017-11-24
 * Copyright: Chris White (www.monitor-computing.co.uk)
 * License:
 **************************************************************/

#include "MainFrame.h"

#include "App.h"

#include "wx_pch.h"
#include <wx/aboutdlg.h>

//(*InternalHeaders(MainFrame)
#include <wx/xrc/xmlres.h>
//*)


//helper functions
enum wxbuildinfoformat {short_f,  long_f };

wxString wxbuildinfo(wxbuildinfoformat format)
{
  wxString wxbuild(wxVERSION_STRING);

  if (format == long_f )
  {
#if defined(__WXMSW__)
    wxbuild << _T("-Windows");
#elif defined(__UNIX__)
    wxbuild << _T("-Linux");
#endif

#if wxUSE_UNICODE
    wxbuild << _T("-Unicode build");
#else
    wxbuild << _T("-ANSI build");
#endif // wxUSE_UNICODE
  }

  return wxbuild;
}

DECLARE_APP(App);


//(*IdInit(MainFrame)
//*)


BEGIN_EVENT_TABLE(MainFrame,wxFrame)
//(*EventTable(MainFrame)
//*)
END_EVENT_TABLE()


MainFrame::MainFrame(wxWindow* parent, wxWindowID id)
{
  //(*Initialize(MainFrame)
  wxXmlResource::Get()->LoadObject(this,parent,_T("MainFrame"),_T("wxFrame"));
  StatusBar1 = (wxStatusBar*)FindWindow(XRCID("cmMainStatusBar"));

  Connect(XRCID("cmLoadConfiguration"),wxEVT_COMMAND_MENU_SELECTED,(wxObjectEventFunction)&MainFrame::on_load_codes);
  Connect(XRCID("cmExit"),wxEVT_COMMAND_MENU_SELECTED,(wxObjectEventFunction)&MainFrame::on_quit);
  Connect(XRCID("cmAbout"),wxEVT_COMMAND_MENU_SELECTED,(wxObjectEventFunction)&MainFrame::on_about);
  //*)
}

MainFrame::~MainFrame()
{
  //(*Destroy(MainFrame)
  //*)
}

void MainFrame::on_quit(wxCommandEvent& event)
{
  Close();
}

void MainFrame::on_load_configuration(wxCommandEvent& event)
{
  wxGetApp().load_configuration_file();
}

void MainFrame::on_about(wxCommandEvent& event)
{
  wxAboutDialogInfo info;
  info.SetName(_("SemaSet"));
  info.SetDescription(_("Servo4/Sema4 configuration utility"));
  info.SetCopyright(wxT("(C) Monitor Computing Services Limited 2019"));

  wxAboutBox(info);
}
