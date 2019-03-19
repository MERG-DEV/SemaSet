/***************************************************************
 * Name:      App.cpp
 * Purpose:   Code for Application Class
 * Author:    Chris White (whitecf@bcs.org.uk)
 **************************************************************/

#include <wx/config.h>
#include <wx/filedlg.h>
#include <wx/filename.h>
#include <wx/msgdlg.h>
#include <wx/string.h>

//(*AppHeaders
#include "MainFrame.h"
#include <wx/xrc/xmlres.h>
#include <wx/image.h>
//*)

#include "wx_pch.h"

#include "App.h"


IMPLEMENT_APP(App);


App::App() :
  m_config{"SemaSet"},
  m_configuration_dir_config_key{"SM4 Directory"}
{
  m_config.Read(m_configuration_dir_config_key, &m_configuration_dir);
}



App::~App()
{
  m_config.Write(m_configuration_dir_config_key, m_configuration_dir);
}



bool
App::OnInit()
{
  //(*AppInitialize
  bool wxsOK = true;
  wxInitAllImageHandlers();
  wxXmlResource::Get()->InitAllHandlers();
  wxsOK = wxsOK && wxXmlResource::Get()->Load(_T("MainFrame.xrc"));
  if ( wxsOK )
  {
  	MainFrame* Frame = new MainFrame(0);
  	Frame->Show();
  	SetTopWindow(Frame);
  }
  //*)

  return wxsOK;
}



void
App::load_configuration_file()
{
  wxFileDialog  file_dialog{GetTopWindow(),
                            _("Open configuration file"),
                            m_configuration_dir, "",
                            _("Sema4 configuration file (*.sm4)|*.sm4"),
                            (wxFD_OPEN|wxFD_FILE_MUST_EXIST)};

  if (wxID_OK == file_dialog.ShowModal())
  {
    m_configuration_dir = wxFileName{file_dialog.GetPath()}.GetPath();
  }
}
