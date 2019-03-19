#pragma once
/***************************************************************
 * Name:      MainFrame.h
 * Purpose:   Defines Application Frame
 * Author:    Chris White (whitecf@bcs.org.uk)
 **************************************************************/

//(*Headers(MainFrame)
#include <wx/frame.h>
#include <wx/menu.h>
#include <wx/statusbr.h>
//*)

class MainFrame: public wxFrame
{
 public:

  MainFrame(wxWindow* parent, wxWindowID id = -1);
  virtual ~MainFrame();

 private:

  //(*Handlers(MainFrame)
  void on_quit(wxCommandEvent& event);
  void on_load_configuration(wxCommandEvent& event);
  void on_about(wxCommandEvent& event);
  //*)

  //(*Identifiers(MainFrame)
  //*)

  //(*Declarations(MainFrame)
  wxMenu* Menu1;
  wxMenu* Menu2;
  wxMenuBar* MenuBar1;
  wxStatusBar* StatusBar1;
  //*)

  DECLARE_EVENT_TABLE()
};
