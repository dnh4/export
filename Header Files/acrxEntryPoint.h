/**
 *  @file acrxEntryPoint.h
 *  @author Romain LEMETTAIS
 *  @brief Point d'entrée du GRX
 */


//-----------------------------------------------------------------------------------<Include>
#pragma once
#include "cmdExport.h"


//-----------------------------------------------------------------------------------<Declaration>
const AcString sGroupName = "TEMPLATE";


//-----------------------------------------------------------------------------------<Fonctions>
/**
  * \brief Decharge l'application
  */
void unloadApp();

/**
  * \brief Initialise l'application
  */
void initApp();