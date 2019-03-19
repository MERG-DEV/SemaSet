#pragma once
/***************************************************************
 * Name:      App.h
 * Purpose:   Defines Application Class
 * Author:    Chris White (whitecf@bcs.org.uk)
 **************************************************************/

#include <wx/app.h>
#include <wx/config.h>
#include <wx/mediactrl.h>
#include <wx/string.h>

class App : public wxApp
{
public:
  App();

  virtual ~App();

  bool
  OnInit() override;

  void
  load_configuration_file();


private:
  wxConfig  m_config;

  wxString  m_configuration_dir_config_key;
  wxString  m_configuration_dir;
};
