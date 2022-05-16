#pragma once
#include "cmdExport.h"
#include "export.h"
#include "poly3DEntity.h"
#include "poly2DEntity.h"
#include "blockEntity.h"
#include "print.h"
#include "layer.h"
#include "file.h"
#include <fstream>
#include <xlnt/xlnt.hpp>
#include <set>
#include <algorithm>

using namespace xlnt;

void cmdExportPoly3d()
{
    //Nom de la polyligne 3d
    ads_name ssPoly3d;
    
    //Recuperer le  chemin du dwg courant
    AcString current_folder = getCurrentFileFolder();
    current_folder += _T( "\\" );
    
    // Le fichier
    AcString file,
             // L'extension du fichier
             ext;
             
    // Tableau de stockage des informations de l'entité
    AcStringArray pData;
    
    //Message : selection
    print( "Veuillez séléctionner la/les polyligne(s) 3D : " );
    
    long size = getSsPoly3D( ssPoly3d, "", false );
    
    //Vérification si il existe des polylignes séléctionnées
    if( size != 0 )
    {
    
        //Nombre d'objet séléctionnés
        int obj = 0;
        
        //2 - Demander à l'utilisateur le répertoire, le nom et l'extension du fichier ou il va enregistrer dans un dossier
        file = askForFilePath( false,
                "xlsx;shp;csv;txt",
                "Enregistrer sous",
                current_folder );
                
        //Gestion d'erreur
        if( file == _T( "" ) )
        {
            print( "Commande annulée." );
            acedSSFree( ssPoly3d );
            return;
        }
        
        //Prendre l'extension du fichier
        ext = getFileExt( file );
        
        //Change file path encodage
        auto filename = latin1_to_utf8( acStrToStr( file ) );
        
        
        //L'entité polyligne 3d
        AcDb3dPolyline* pl3d = NULL;
        
        /*
            --------------------------------------------
            -->Si le fichier est un fichier .csv ou .txt
            --------------------------------------------
        */
        if( ext == "csv" || ext == "txt" )
        {
            //Ouvverture du fichier obtenue dans askForFilePath()
            fstream fichier( acStrToStr( file ), ios::app );
            
            //Ecriture des en-têtes du fichier
            fichier << "Handle;Calque;Couleur;Type de ligne;Echelle de type de ligne;Epaisseur de ligne;Transparence;Materiau;Longueur 2d;Longueur 3d;Fermée" << endl;
            
            //3 - Récupérer les informations sur la/les polyligne(s)
            ProgressBar prog_txt = ProgressBar( _T( "Progression:" ), size );
            
            // Iteration sur les objets
            for( long i = 0; i < size; i++ )
            {
            
                //Le pointeur sur la polyligne
                pl3d = getPoly3DFromSs( ssPoly3d, i, GcDb::kForRead );
                
                //verifier pl3d
                if( !pl3d )
                {
                    //Progresser
                    prog_txt.moveUp( i );
                    
                    continue;
                }
                
                // Appel de la fonction de traitement
                exportPoly3d( pl3d, pData );
                
                
                //Exclure la surface
                pData.removeLast();
                
                // Iteration sur le tableau de données
                for( long j = 0; j < pData.size(); j++ )
                {
                    //Ecriture des données dans le fichier
                    fichier << pData[j] << ";";
                }
                
                // Se mettre à la ligne
                fichier << endl;
                
                //Incrémentation du nombre d'objet
                obj++;
                
                // Vider le tableau des données
                pData.clear();
                
                //Fermeture de la polyligne
                pl3d->close();
                
                prog_txt.moveUp( i );
            }
            
            try
            {
                //Fermeture du fichier
                fichier.close();
                
                //Affichage dans le console que les informations sont bien exportés
                print( "Exportation de " + to_string( obj ) + " polylignes 3D dans " + acStrToStr( file ) + " terminée." );
            }
            
            catch( ... )
            {
                //message
                print( "Impossible d'enregistrer les modifications." );
                
                //Liberer la mémoire
                acedSSFree( ssPoly3d );
            }
        }
        
        /*
            ---------------------------------------------
            -->Si le fichier est un fichier .xlsx
            ---------------------------------------------
        */
        else if( ext == "xlsx" )
        {
            //Initialisation du fichier excel
            workbook wb2;
            worksheet ws = wb2.active_sheet();
            
            //Ajout des en-tête du feuille
            ws.cell( 1, 1 ).value( "Handle" );
            ws.cell( 2, 1 ).value( "Calque" );
            ws.cell( 3, 1 ).value( "Couleur" );
            ws.cell( 4, 1 ).value( "Type de ligne" );
            ws.cell( 5, 1 ).value( "Echelle de type de ligne" );
            ws.cell( 6, 1 ).value( "Epaisseur de ligne" );
            ws.cell( 7, 1 ).value( "Transparence" );
            ws.cell( 8, 1 ).value( "Materiau" );
            ws.cell( 9, 1 ).value( "Longueur 2d" );
            ws.cell( 10, 1 ).value( "Longueur 3d" );
            ws.cell( 11, 1 ).value( "Fermee" );
            
            //Barre de progression
            ProgressBar prog_xls = ProgressBar( _T( "Progression: " ), size );
            
            for( long i = 0; i < size; i++ )
            {
            
                //Le pointeur sur la polyligne
                pl3d = getPoly3DFromSs( ssPoly3d, i, GcDb::kForRead );
                
                //Safeguard
                if( !pl3d )
                {
                    //Progresser
                    prog_xls.moveUp( i );
                    
                    continue;
                }
                
                // Appel de la fonction de traitement
                exportPoly3d( pl3d, pData );
                
                // Mettre le format de toute les colonnes en texte
                ws.columns( true ).number_format( number_format::text() );
                
                //Exclure la surface
                pData.removeLast();
                
                // Iteration sur le tableau des données
                for( long k = 0; k < pData.size(); k++ )
                {
                    //Convertir l'encodage des strings
                    auto data = latin1_to_utf8( acStrToStr( pData[k] ) );
                    
                    //Ecriture des données dans le fichier excel
                    ws.cell( k + 1, i + 2 ).value( data );
                }
                
                //Incrémentation du nombre d'objet
                obj++;
                
                // Vider le tableau des données
                pData.clear();
                
                //Fermer la polyligne
                pl3d->close();
                
                //Progresser
                prog_xls.moveUp( i );
            }
            
            try
            {
                wb2.save( filename );
                
                //Affichage dans la console que les informations sont bien exportés
                print( "Exportation de " + to_string( obj ) + " polyligne 3D dans " + acStrToStr( file ) + " terminée." );
            }
            
            catch( ... )
            {
                //Message
                print( "Impossible d'enregistrer les modifications." );
                
                //Liberer la mémoire
                acedSSFree( ssPoly3d );
            }
            
            
            
        }
        
        /*
            --------------------------------------------
            -->Si le fichier est un fichier .shp
            --------------------------------------------
        
        */
        
        else if( ext == "shp" )
        {
        
            //Lire le fichier de parametre
            vector<AcString> field = createPoly3dField();
            writePrjOrNot( file );
            
            //Creer un fichier shp
            SHPHandle myShapeFile = SHPCreate( file, SHPT_ARCZ );
            
            //Creer un dbfFile
            DBFHandle dbfHandle = DBFCreate( file );
            
            //Creer les champs
            int fieldNumber = createFields( dbfHandle, field );
            
            
            // Initialistion de la barre de progression
            ProgressBar prog = ProgressBar( _T( "Progression" ), size );
            
            //Nombre d'objet exportés
            int nb_el = 0;
            
            //Boucler sur la selection
            for( int i = 0; i < size; i++ )
            {
            
                //Recuperer l'entite
                pl3d = getPoly3DFromSs( ssPoly3d, i );
                
                //Verifier la polyline
                if( !pl3d )
                    continue;
                    
                //Reciperer l'objet ligne
                drawLineShp( pl3d, myShapeFile, dbfHandle, field );
                
                //Fermer la polyline
                pl3d->close();
                
                // incrementer la barre de progression
                prog.moveUp( i );
                
                //Incrementer le nombre d'objet exportés
                nb_el++;
            }
            
            //Liberer ma mémoire
            DBFClose( dbfHandle );
            SHPClose( myShapeFile );
            
            //Affichage dans la console que les informations sont bien exportés
            print( "Exportation de " + to_string( nb_el ) + " polyligne 3D dans " + acStrToStr( file ) + " terminée." );
            
        }
        
    }
    
    //Si il n'y a pas de polyline séléctionnée
    else
        print( "Aucune polyligne séléctionnée ! " );
        
    //Liberer la selection
    acedSSFree( ssPoly3d );
}

void cmdExportPoly2d()
{
    // Nom de la polyligne 2d
    ads_name ssPoly2d;
    
    // Le fichier
    AcString file,
             // L'extension du fichier
             ext;
             
    //Initialisation
    file = _T( "" );
    
    
    //Recuperer le  chemin du dwg courant
    AcString current_folder = getCurrentFileFolder();
    current_folder += _T( "\\" );
    
    // Tableau de stockage des données des entités
    AcStringArray pData;
    
    // Demander à l'utilisateur de séléctionner les polylignes
    print( "Veuillez sélectionner les polylignes : " );
    
    // Les sélection sur les polylignes
    long size = getSsPoly2D( ssPoly2d, _T( "" ), false );
    
    // Verification si la sélection n'est pas vide
    if( size != 0 )
    {
        // Nombre d'objet sélectionnées
        int obj = 0;
        
        // Demander à l'utilisateur le repertoire et l'extension du fichier à enregistrer
        file = askForFilePath( false,
                "xlsx;shp;csv;txt",
                "Enregistrer sous",
                current_folder );
                
        // Prendre l'extension fu fichier
        ext = getFileExt( file );
        
        //Changer l'encodage du chemin pour le fichier excel
        auto filename = latin1_to_utf8( acStrToStr( file ) );
        
        // L'entité polyligne 2d
        AcDbPolyline* pl2d = NULL;
        
        // Si le fichier est .txt ou .csv
        if( ext == "txt" || ext == "csv" )
        {
            // Flux de fichier
            fstream fichier( acStrToStr( file ), ios::app );
            
            // Ecriture des en-têtes du fichier
            fichier << "Handle;Calque;Couleur;Type de ligne;Echelle de type de ligne;Epaisseur de ligne;Transparence;Epaisseur;Materiau;Longueur 2d;Largeur;Fermée" << endl;
            
            ProgressBar prog_txt = ProgressBar( _T( "Progression:" ), size );
            
            // Iteration sur les objets
            for( long i = 0; i < size; i++ )
            {
                // Le pointeur sur la polyligne
                pl2d = getPoly2DFromSs( ssPoly2d, i, AcDb::kForRead );
                
                //Verifier pl2d
                if( !pl2d )
                {
                    prog_txt.moveUp( i );
                    
                    continue;
                }
                
                // Appel de la fonction de traitement
                exportPoly2d( pl2d, pData );
                
                //Exclure la surface
                pData.removeLast();
                
                // Iteration sur le tableau des données
                for( long j = 0; j < pData.size(); j++ )
                {
                    // Ecriture des données dans le fichier
                    fichier << pData[j] << ";";
                }
                
                // Se mettre à la ligne
                fichier << endl;
                
                // Incrémentation du nomber d'objet
                obj++;
                
                // Vider le tableau des données
                pData.clear();
                
                // Fermer la polyligne
                pl2d->close();
                prog_txt.moveUp( i );
            }
            
            try
            {
                // Fermer le flux de fichier
                fichier.close();
                
                // Affichage dans la console que les informations sont bien exportés
                print( "Exportation de " + to_string( obj ) + " polyligne(s) 2D dans " + acStrToStr( file ) + " terminée." );
            }
            
            catch( ... )
            {
                print( "Impossible d'enregistrer le fichier." );
                acedSSFree( ssPoly2d );
            }
        }
        
        // Si le fichier est .xlsx(excel)
        else if( ext == "xlsx" )
        {
            // Initialisation du fichier excel
            workbook wb;
            worksheet ws = wb.active_sheet();
            
            // Ajout des en-têtes dans le feuille excel
            ws.cell( 1, 1 ).value( "Handle" );
            ws.cell( 2, 1 ).value( "Calque" );
            ws.cell( 3, 1 ).value( "Couleur" );
            ws.cell( 4, 1 ).value( "Type de ligne" );
            ws.cell( 5, 1 ).value( "Echelle de type de ligne" );
            ws.cell( 6, 1 ).value( "Epaisseur de ligne" );
            ws.cell( 7, 1 ).value( "Transparence" );
            ws.cell( 8, 1 ).value( "Epaisseur" );
            ws.cell( 9, 1 ).value( "Materiau" );
            ws.cell( 10, 1 ).value( "Longueur 2d" );
            ws.cell( 11, 1 ).value( "Largeur" );
            ws.cell( 12, 1 ).value( "Fermee" );
            
            ProgressBar prog_xls = ProgressBar( _T( "Progression:" ), size );
            
            //Compteur de ligne
            int row_count = 0;
            
            // Iteration sur les objets
            for( long i = 0; i < size; i++ )
            {
                // Le pointeur sur la polyligne
                pl2d = getPoly2DFromSs( ssPoly2d, i, AcDb::kForRead );
                
                // Safeguard
                if( !pl2d )
                {
                    //progresser
                    prog_xls.moveUp( i );
                    continue;
                }
                
                // Appel de la fonction de traitement
                exportPoly2d( pl2d, pData );
                
                //Exclure la surface
                pData.removeLast();
                
                // Mettre le format de toute les cellules en Texte
                ws.columns( true ).number_format( number_format::text() );
                
                // Iteration sur le tableau des données
                for( long j = 0; j < pData.size(); j++ )
                {
                    //Convertir l'encodage des strings
                    auto data = latin1_to_utf8( acStrToStr( pData[j] ) );
                    
                    // Ecriture des données dans le fichier
                    ws.cell( j + 1, row_count + 2 ).value( data );
                }
                
                // Incrémentation du nombre d'objet
                obj++;
                
                //Incrementer la ligne
                row_count++;
                
                // Vider le tableau des données
                pData.clear();
                
                // Fermer la polyligne 2d
                pl2d->close();
                
                //Progresser
                prog_xls.moveUp( i );
            }
            
            try
            {
                // Sauvegarder le fichier excel .xlsx
                wb.save( filename );
                
                // Affichage dans la console que les propriétes sont bien exportés
                print( "Exportation de " + to_string( obj ) + " polyligne(s) 2D dans " + acStrToStr( file ) + " terminée." );
                
            }
            
            catch( ... )
            {
                print( "Impossible d'enregistrer le fichier." );
                acedSSFree( ssPoly2d );
            }
            
        }
        
        //Si le fichier est .shp(shape)
        
        else if( ext == "shp" )
        {
        
        
            //Lire le fichier de parametre
            vector<AcString> field = createPolyField();
            writePrjOrNot( file );
            
            //Creer un fichier shp
            SHPHandle myShapeFile = SHPCreate( file, SHPT_ARCZ );
            
            //Creer un dbfFile
            DBFHandle dbfHandle = DBFCreate( file );
            
            //Creer les champs
            int fieldNumber = createFields( dbfHandle, field );
            
            
            // Initialistion de la barre de progression
            ProgressBar prog = ProgressBar( _T( "Progression" ), size );
            
            //Boucler sur la selection
            for( int i = 0; i < size; i++ )
            {
                int vertNumber = 0;
                //Recuperer l'entite
                pl2d = getPoly2DFromSs( ssPoly2d, i );
                
                //Verifier ent
                if( !pl2d )
                {
                    // incrementer la barre de progression
                    prog.moveUp( i );
                    
                    continue;
                }
                
                //Reciperer l'objet ligne
                drawLineShp( pl2d, myShapeFile, dbfHandle, field );
                
                //Liberer la mémoire
                pl2d->close();
                
                // incrementer la barre de progression
                prog.moveUp( i );
            }
            
            DBFClose( dbfHandle );
            SHPClose( myShapeFile );
            
            // Affichage dans la console que les propriétes sont bien exportés
            print( "Exportation de " + to_string( size ) + " polyligne(s) 2D dans " + acStrToStr( file ) + " terminée." );
            
        }
        
    }
    
    // Si la sélection est vide
    else
    {
        // Affichage dans la console
        print( "Aucune polyligne sélectionnée." );
    }
    
    // Liberer la sélection
    acedSSFree( ssPoly2d );
    
}

void cmdExportPoint()
{
    //Selection des points
    ads_name ssPoint;
    
    //Fichier & extension
    AcString file, ext;
    
    //Recuperer le  chemin du dwg courant
    AcString current_folder = getCurrentFileFolder();
    current_folder += _T( "\\" );
    
    // Tableau de stockage des données
    AcStringArray pData;
    
    //Demander au dessinateur de faire une selection
    print( "Veuillez séléctionner le/les point(s) : " );
    
    //Nombre de points dans la selection
    long size = getSsPoint( ssPoint, "" );
    
    //Verifier le nombre de points
    if( size != 0 )
    {
    
        //Nombre d'objet sélectionné
        int obj = 0;
        
        //Initialiser file
        file = _T( "" );
        
        //Ou enregistrer le fichier?
        file = askForFilePath( false,
                "xlsx;shp;csv;txt",
                "Enregistrer sous",
                current_folder );
                
        //Recuperer l'extension du fichier
        ext = getFileExt( file );
        
        //Gestion d'erreur
        if( file == _T( "" ) )
        {
            print( "Commande annulée." );
            acedSSFree( ssPoint );
            return;
        }
        
        //Changer l'encodage du chemin vers le fichier excel en utf8
        auto filename = latin1_to_utf8( acStrToStr( file ) );
        
        //Déclarer un AcDbPoint
        AcDbPoint* pt = NULL;
        
        //Si : .txt, .csv
        if( ext == "csv" || ext == "txt" )
        {
            //Ouvrir le fichier
            fstream fichier( acStrToStr( file ), ios::app );
            
            //Ecrire les en-têtes du fichier
            fichier << "Handle;Calque;Couleur;Type de ligne;Echelle de type de ligne;Epaisseur de ligne;Transparence;Epaisseur;Materiau;X;Y;Z" << endl;
            
            //Barre de progression
            ProgressBar prog_txt = ProgressBar( _T( "Progression:" ), size );
            
            //Iterer sur la selection de points
            for( long i = 0; i < size; i++ )
            {
            
                //Le pointeur sur le point
                pt = getPointFromSs( ssPoint, i, GcDb::kForRead );
                
                //Safeguard
                if( !pt )
                {
                    prog_txt.moveUp( i );
                    continue;
                }
                
                // Appel de la fonction de traitement
                exportPoint( pt, pData );
                
                // Iteration sur le tableau des données
                for( long j = 0; j < pData.size(); j++ )
                {
                    //Ecriture des données dans le fichier sélectionné
                    fichier << pData[j] << ";";
                }
                
                // Se mettre à la ligne
                fichier << endl;
                
                //Incrémentation du nombre d'objet
                obj++;
                
                //Reset pData
                pData.clear();
                
                //Liberer la mémoire
                pt->close();
                
                //Progresser
                prog_txt.moveUp( i );
            }
            
            try
            {
                //Fermeture du fichier
                fichier.close();
                
                //Affichage dans la console que les informations sont bien exportés
                print( "Exportation de " + to_string( obj ) + " points dans " + acStrToStr( file ) + " terminée." );
            }
            
            catch( ... )
            {
                print( "Impossible d'enregistrer le fichier." );
                acedSSFree( ssPoint );
            }
        }
        
        //Fichier excel
        else if( ext == "xlsx" )
        {
            //Initialisation du fichier excel
            workbook wb2;
            worksheet ws = wb2.active_sheet();
            
            //Ajout des en-tête du feuille
            ws.cell( 1, 1 ).value( "Handle" );
            ws.cell( 2, 1 ).value( "Calque" );
            ws.cell( 3, 1 ).value( "Couleur" );
            ws.cell( 4, 1 ).value( "Type de ligne" );
            ws.cell( 5, 1 ).value( "Echelle de type de ligne" );
            ws.cell( 6, 1 ).value( "Epaisseur de ligne" );
            ws.cell( 7, 1 ).value( "Transparence" );
            ws.cell( 8, 1 ).value( "Epaisseur" );
            ws.cell( 9, 1 ).value( "Materiau" );
            ws.cell( 10, 1 ).value( "X" );
            ws.cell( 11, 1 ).value( "Y" );
            ws.cell( 12, 1 ).value( "Z" );
            
            //BArre de progression
            ProgressBar prog = ProgressBar( _T( "Progression" ), size );
            
            //Compteur de ligne
            int row_count = 0;
            
            for( long i = 0; i < size; i++ )
            {
                //Recuperer le i-eme point
                pt = getPointFromSs( ssPoint, i, GcDb::kForRead );
                
                //Safeguard
                if( !pt )
                {
                    prog.moveUp( i );
                    continue;
                }
                
                // Appel de la fonction de traitement
                exportPoint( pt, pData );
                
                // Mettre le format des cellules en texte
                ws.columns( true ).number_format( number_format::text() );
                
                // Iteration sur les données
                for( long k = 0; k < pData.size(); k++ )
                {
                    //Ecriture dans le fichier excel
                    ws.cell( k + 1, row_count + 2 ).value( pData[k] );
                }
                
                //Incrémentation du nombre d'objet
                obj++;
                
                //Incrementer le nombre de lignes
                row_count++;
                
                // Vider le tableau des données
                pData.clear();
                
                //Fermer le point
                pt->close();
                
                //Progresser
                prog.moveUp( i );
            }
            
            try
            {
                //Enregistrer les modifications
                wb2.save( filename );
                
                //Affichage dans la console que les informations sont bien exportés
                print( "Exportation de " + to_string( obj ) + " points dans " + acStrToStr( file ) + " terminée." );
                
            }
            
            catch( ... )
            {
                print( "Impossible d'enregistrer le fichier." );
                acedSSFree( ssPoint );
            }
        }
        
        //Fichier shape
        else if( ext == "shp" )
        {
        
            //Lire le fichier de parametre
            vector<AcString> field = createPointField();
            writePrjOrNot( file );
            
            //Creer un fichier shp
            SHPHandle myShapeFile = SHPCreate( file, SHPT_POINTZ );
            
            //Creer un dbfFile
            DBFHandle dbfHandle = DBFCreate( file );
            
            //Creer les champs
            int fieldNumber = createFields( dbfHandle, field );
            
            // Initialistion de la barre de progression
            ProgressBar prog = ProgressBar( _T( "Progression" ), size );
            
            //Nombre de points exportés
            int nb_el = 0;
            
            //Boucler sur toute les blocs
            for( int i = 0; i < size; i++ )
            {
                //Recuperer le i-eme point
                pt = getPointFromSs( ssPoint, i );
                
                //Verifier pt
                if( !pt )
                {
                    prog.moveUp( i );
                    continue;
                }
                
                //Exporter les points
                drawPointShp( pt, myShapeFile, dbfHandle, field );
                
                //Liberer la mémoire
                pt->close();
                
                //Incrementer le nombre de points éxportés
                nb_el++;
                
                //Progresser
                prog.moveUp( i );
            }
            
            
            //Liberer la mémoire
            DBFClose( dbfHandle );
            SHPClose( myShapeFile );
            
            // Messages
            print( "Exportation de " + to_string( nb_el ) + " points dans " + acStrToStr( file ) + " terminée." );
            
        }
    }
    
    //Si il n'y a pas de point sélectionné
    else
        print( "Aucun point sélectionné" );
        
        
    acedSSFree( ssPoint );
}

void cmdExportBlock()
{

    //Définition de la selection
    ads_name ssBlock;
    ads_name ssCopie;
    
    //Fichier & Extension
    AcString file, ext;
    
    file = _T( "" );
    
    //Recuperer le  chemin du dwg courant
    AcString current_folder = getCurrentFileFolder();
    current_folder += _T( "\\" );
    
    //Stockage : propriétés & données
    AcStringArray pData;
    AcStringArray propName;
    
    //Vecteur temporaire: liste des attributs & liste des noms de propriete
    vector<AcString> nomAttributs, nomPropriete;
    
    //Vecteur de type non fixe pour contenir les valeurs des proprietes
    vector<void*> propValue;
    
    //Set : (Noms des attributs sans doublons & Noms des propriétes sans doublons & Noms de définition) pour toutes les entités
    set<AcString> sAttr, sProp;
    
    //Nombre de blocs dynamiques
    int dynBlocCount = 0;
    
    //Message
    print( "Veuillez sélectionner le/les block(s) : " );
    
    //Recuperer le nombre de blockes
    long size = getSsBlock( ssBlock, "", "" );
    
    //Verifier
    if( size != 0 )
    {
    
        //Nombre d'objet exporté
        int obj = 0;
        
        //Demander chemin du fichier vers quoi on exporte
        file = askForFilePath( false,
                "xlsx;shp;csv;txt",
                "Enregistrer sous",
                current_folder );
                
        //Gestion d'erreur
        if( file == _T( "" ) )
        {
            print( "Commande annulée." );
            acedSSFree( ssBlock );
            return;
        }
        
        //Recuperer l'extension
        ext = getFileExt( file );
        
        //Changer l'encodage du chemin vers le fichier excel en utf8
        string fname = acStrToStr( file );
        
        auto filename = latin1_to_utf8( fname );
        
        //Déclarer une reference de blocke
        AcDbBlockReference* blk;
        
        //Fichier : .txt ou .csv
        if( ext == "txt" || ext == "csv" )
        {
            //Ouvrir le fichier
            fstream fichier( acStrToStr( file ), ios::app );
            
            //Ecriture des en-têtes du fichier
            fichier << "Handle;Calque;Nom;Couleur;Type de ligne;Echelle de type de ligne;Epaisseur de ligne;Transparence;Materiau;Pos_X;Pox_Y;Pos_Z;Rotation";
            
            //Itération sur les objets
            for( long i = 0; i < size; i++ )
            {
                //Pointeur sur le bloc
                blk = getBlockFromSs( ssBlock, i );
                
                //Recuperer la taille
                int taille_attr = getAttributesNamesListOfBlockRef( blk ).size();
                
                //Iteration sur la taille de la liste des noms des attributs
                for( long j = 0; j < taille_attr; j++ )
                {
                    //Mettre les attributs dans le set
                    sAttr.insert( getAttributesNamesListOfBlockRef( blk )[j] );
                }
                
                //Si block dynamique
                if( isDynamicBlock( blk ) )
                {
                    //Inrementation du nombre de bloc dynamique
                    dynBlocCount++;
                    
                    //Prendre les listes des noms de proprietes et aussi des valeurs de proprietes
                    getBlockPropWithValueList( blk, propName, propValue );
                    
                    //Iteration sur la liste des noms des proprietes
                    for( long pn = 0; pn < propName.size(); pn++ )
                    {
                        //Mettre les noms des proprietes dans le set (sans doublons)
                        sProp.insert( propName[pn] );
                    }
                }
                
                //Fermeture du bloc
                blk->close();
            }
            
            //Itération sur les données sans doublons dans le |setAttr| et les mettre dans le vecteur temporaire
            for( set<AcString>::iterator t = sAttr.begin(); t != sAttr.end(); t++ )
            {
                nomAttributs.push_back( *t );
                //Completer les en-tête avec la liste sans doublons des ATTRIBUTS
                fichier << ";" << acStrToStr( *t );
            }
            
            //Iteration sur les données sans doublons dans le |sProp| et les mettre dans le vecteur temporaire
            for( set<AcString>::iterator r = sProp.begin(); r != sProp.end(); r++ )
            {
                //Mettre les noms de proprietés sans doublons dans un vecteur
                nomPropriete.push_back( *r );
                //Completer les en-têtes avec la liste sans doublons des PROPRIETES
                fichier << ";" << acStrToStr( *r );
            }
            
            //Inserer l'entete nombre de copies
            fichier << ";Nombre de copies";
            
            //Se mettre à la ligne pour commencer à inserer les données
            fichier << endl;
            
            //Barre de progression
            ProgressBar prog = ProgressBar( _T( "Progression :" ), size );
            
            //Reiterer sur les blocks
            for( long i = 0; i < size; i++ )
            {
            
                //Pointeur sur le block
                blk = getBlockFromSs( ssBlock, i );
                
                // Appel de la fonction de traitement
                exportBlock( blk, pData );
                
                // Iteration sur les données
                for( long j = 0; j < pData.size(); j++ )
                {
                    //Ecrire les données
                    fichier << pData[j] << ";";
                }
                
                //Rcuperer la taille de nomAttributs
                int nb_att = nomAttributs.size();
                
                //Itération sur les attributs
                for( long att = 0; att < nb_att; att++ )
                {
                    // Si la valeur de l'attribut n'est pas vide
                    if( getAttributValue( blk, nomAttributs[att] ) != "empty" )
                    {
                        //Convertir en std::string
                        auto attrib_val = acStrToStr( getAttributValue( blk, nomAttributs[att] ) );
                        
                        //Ecrire la valeur de l'attribut
                        fichier << attrib_val << ";";
                    }
                    
                    // Si la valeur de l'attribut est vide
                    else
                    {
                        // Remplir la colonne par des N/A
                        fichier << "N/A;";
                    }
                }
                
                //Check si c'est un bloc dynamique
                if( isDynamicBlock( blk ) )
                {
                
                    //Iteration sur les noms de proprieté sans doublons
                    for( long propName = 0; propName < nomPropriete.size(); propName++ )
                    {
                        //Recuperer la valeur du propriete
                        AcString bProp = getPropertyStrFromBlock( blk, nomPropriete[propName] );
                        
                        //Si bProp n'est pas vide
                        if( bProp != _T( "" ) )
                        {
                            //Recuperer la valeur du propriete
                            auto prop_val = acStrToStr( getPropertyStrFromBlock( blk, nomPropriete[propName] ) );
                            
                            //Ecrire dans le fichier
                            fichier << prop_val << ";";
                        }
                        
                        // Si la valeur du propriete n'est pas vide
                        else
                        {
                            // Si la valeur du propriete est un double
                            if( getPropertyFromBlock( blk, nomPropriete[propName] ) != -1 || getPropertyFromBlock( blk, nomPropriete[propName] ) != 0 )
                            {
                                //Recuperer la propriete
                                auto prop_value = to_string( getPropertyFromBlock( blk, nomPropriete[propName] ) );
                                
                                //Ecrire les valeurs des proprietés dans le fichier
                                fichier << prop_value << ";";
                            }
                            
                            // Si ce n'est pas les deux types
                            else
                                fichier << "N/A;";
                        }
                        
                    }
                    
                }
                
                // Si ce n'est pas un bloc dynamique
                else
                {
                    for( long propName = 0; propName < nomPropriete.size(); propName++ )
                        fichier << "N/A;";
                }
                
                //Se mettre à la ligne après une écriture d'une ligne d'information des objets
                fichier << endl;
                
                //Incrémentation de obj
                obj++;
                
                //Fermeture du block
                blk->close();
                
                
                // Vider le tableau des données
                pData.clear();
                
                //Progresser
                prog.moveUp( i );
            }
            
            try
            {
                //Fermeture du fichier
                fichier.close();
                
                //Affichage dans la console que les informations sont bien exportés
                print( "Exportation de " + to_string( obj ) + " block(s) dont " + to_string( dynBlocCount ) + " bloc(s) dynamique(s) dans " + acStrToStr( file ) + " terminée." );
                
            }
            
            catch( ... )
            {
                print( "Impossible d'enregistrer le fichier." );
                acedSSFree( ssBlock );
            }
            
        }
        
        // Si c'est un fichier .xlsx
        else if( ext == "xlsx" )
        {
        
            //Initialisation du fichier excel
            workbook wb;
            worksheet ws = wb.active_sheet();
            
            //----------------------------------------------------------------------------------------------------<Debut en-tête>
            
            //Ajout des en-tête
            ws.cell( 1, 1 ).value( "Handle" );
            ws.cell( 2, 1 ).value( "Calque" );
            ws.cell( 3, 1 ).value( "Nom" );
            ws.cell( 4, 1 ).value( "Couleur" );
            ws.cell( 5, 1 ).value( "Type de ligne" );
            ws.cell( 6, 1 ).value( "Echelle de type de ligne" );
            ws.cell( 7, 1 ).value( "Epaisseur de ligne" );
            ws.cell( 8, 1 ).value( "Transparence" );
            ws.cell( 9, 1 ).value( "Materiau" );
            ws.cell( 10, 1 ).value( "Pos_X" );
            ws.cell( 11, 1 ).value( "Pos_Y" );
            ws.cell( 12, 1 ).value( "Pos_Z" );
            ws.cell( 13, 1 ).value( "Rotation" );
            
            //Itération sur les objets
            for( long i = 0; i < size; i++ )
            {
            
                //Pointeur sur le bloc
                blk = getBlockFromSs( ssBlock, i );
                
                //Recuperer la taille
                int taille_attr = getAttributesNamesListOfBlockRef( blk ).size();
                
                //Iteration sur la taille de la liste des noms des attributs
                for( long j = 0; j < taille_attr; j++ )
                {
                    //Mettre les attributs dans le set
                    sAttr.insert( getAttributesNamesListOfBlockRef( blk )[j] );
                }
                
                //Si block dynamique
                if( isDynamicBlock( blk ) )
                {
                    //Inrementation du nombre de bloc dynamique
                    dynBlocCount++;
                    
                    //Prendre les listes des noms de proprietes et aussi des valeurs de proprietes
                    getBlockPropWithValueList( blk, propName, propValue );
                    
                    //Iteration sur la liste des noms des proprietes
                    for( long pn = 0; pn < propName.size(); pn++ )
                    {
                        //Mettre les noms des proprietes dans le set (sans doublons)
                        sProp.insert( propName[pn] );
                    }
                }
                
                //Fermeture du bloc
                blk->close();
            }
            
            //Iteration sur les données sans doublons dans le |setAttr| et les mettre dans le vecteur
            for( set<AcString>::iterator t = sAttr.begin(); t != sAttr.end(); t++ )
            {
                // Mettre les données sans doublons dans un vecteur
                nomAttributs.push_back( *t );
            }
            
            //Recuperer la taille de nomAttributs
            int ta = nomAttributs.size();
            
            //Completer les colonnes d'en tête
            for( int a = 0; a < ta; a++ )
            {
                // Completer les données dans le fichier excel
                ws.cell( 13 + ( a + 1 ), 1 ).value( nomAttributs[a] );
            }
            
            //Iteration sur les données sans doublons dans le |sProp| et les mettre dans le vecteur temporaire
            for( set<AcString>::iterator r = sProp.begin(); r != sProp.end(); r++ )
            {
                //Mettre les noms des proprietes sans doublons dans le vecteur
                nomPropriete.push_back( *r );
            }
            
            //Recuperer la taille du nomPropriete
            int tp = nomPropriete.size();
            
            //Recuperer la taille du nomAttributs
            int tt = nomAttributs.size();
            
            //Colonne d'insertion nombre de copies
            int col_nbcop = 0;
            
            //Iteration sur les noms de propriete
            for( int p = 0; p < tp; p++ )
            {
                //Completer les en-têtes dans le fichier excel
                ws.cell( ( 13 + tt + ( p + 1 ) ), 1 ).value( nomPropriete[p] );
                col_nbcop++;
            }
            
            //Barre de progression
            ProgressBar prog_xls = ProgressBar( _T( "Progression :" ), size );
            
            //Indice de ligne
            int ro = 2;
            
            //Iterer sur les blockes
            for( long i = 0; i < size; i++ )
            {
            
                //Recuperer le i-eme block
                blk = getBlockFromSs( ssBlock, i, GcDb::kForRead );
                
                // Definir les formats de toutes les cellules dans les colonnes en texte
                ws.columns( true ).number_format( xlnt::number_format::text() );
                
                //Recuperer les données
                exportBlock( blk, pData );
                
                //Ecrire les données
                for( long d = 0; d < pData.size(); d++ )
                    ws.cell( d + 1, ro ).value( pData[d] );
                    
                //Iteration sur les attributs
                for( long att = 0; att < nomAttributs.size(); att++ )
                {
                    if( getAttributValue( blk, nomAttributs[att] ) != "empty" )
                        ws.cell( 13 + ( att + 1 ), ro ).value( getAttributValue( blk, nomAttributs[att] ) );
                    else
                        ws.cell( 13 + ( att + 1 ), ro ).value( "N/A" );
                }
                
                // Check si c'est un bloc dynamique
                if( isDynamicBlock( blk ) )
                {
                    for( long propName = 0; propName < nomPropriete.size(); propName++ )
                    {
                        /*
                            OBTENIR LES VALEURS DES PROPRIETES A L'AIDE DES NOMS DE PROPRIETES
                        */
                        
                        AcString nProp = getPropertyStrFromBlock( blk, nomPropriete[propName] );
                        
                        // Si la valeur du propriete est un AcString
                        if( nProp != _T( "" ) )
                        {
                            //Recuperer la taille de nomAttributs
                            int tai_n = nomAttributs.size();
                            
                            //Recuperer le nom de propriete
                            auto pro_name = acStrToStr( getPropertyStrFromBlock( blk, nomPropriete[propName] ) );
                            
                            // Completer les colonnes excel avec les valeurs de proprietes
                            ws.cell( ( 13 + tai_n + 1 + propName ), ro ).value( pro_name );
                        }
                        
                        // Si la valeur du propriete n'est pas un AcString
                        else
                        {
                            // Si la valeur est un double
                            if( getPropertyFromBlock( blk, nomPropriete[propName] ) != -1 || getPropertyFromBlock( blk, nomPropriete[propName] ) != 0 )
                            {
                                // Ecrire les valeurs des proprietes dans les colonnes correspondantes
                                ws.cell( ( 13 + nomAttributs.size() + 1 + propName ), ro ).value( to_string( getPropertyFromBlock( blk, nomPropriete[propName] ) ) );
                            }
                            
                            // Si ce n'est pas dans les deux cas
                            else
                            {
                                // Mettre N/A à la place
                                ws.cell( ( 13 + nomAttributs.size() + 1 + propName ), ro ).value( "N/A" );
                            }
                        }
                    }
                }
                
                // Si ce n'est pas un blooc dynamique
                else
                {
                    // Iteration sur les noms de proprietes
                    for( long pNm = 0; pNm < nomPropriete.size(); pNm++ )
                        // Completer par N/A
                        ws.cell( ( 13 + nomAttributs.size() + 1 + pNm ), ro ).value( "N/A" );
                }
                
                //Place nombre copies
                
                //Incrementation de obj
                obj++;
                
                //Fermeture du block
                blk->close();
                
                // Vider le tableau des données
                pData.clear();
                
                //Incrementer la ligne
                ro++;
                
                //Progresser
                prog_xls.moveUp( i );
            }
            
            try
            {
                //Sauvegarder le fichier excel
                wb.save( filename );
                
                //Affichage dans la console le nombre de bloc(s) ou on a exporter ses informations + le nombre de bloc(s) dynamique dans la selection
                print( "Exportation de " + to_string( obj ) + " block(s) dont " + to_string( dynBlocCount ) + " dynamique(s) dans " + acStrToStr( file ) + " terminée." );
                
            }
            
            catch( ... )
            {
                print( "Impossible d'enregistrer le fichier." );
                acedSSFree( ssBlock );
            }
            
        }
        
        //Si c'est un fichier .shp
        else if( ext == "shp" )
        {
            //Lire le fichier de parametre
            vector<AcString> field = createBlockField();
            writePrjOrNot( file );
            
            //Creer un fichier shp
            SHPHandle myShapeFile = SHPCreate( file, SHPT_POINTZ );
            
            //Creer un dbfFile
            DBFHandle dbfHandle = DBFCreate( file );
            
            //Creer les champs
            int fieldNumber = createFields( dbfHandle, field );
            
            // Initialistion de la barre de progression
            ProgressBar prog = ProgressBar( _T( "Progression" ), size );
            
            //Nombre de block exportés
            int nb_el = 0;
            
            //Boucler sur toute les blocs
            for( int i = 0; i < size; i++ )
            {
                //Recuperer l'objet block
                blk = getBlockFromSs( ssBlock, i );
                
                //Verifier blk
                if( !blk )
                {
                    //Progresser
                    prog.moveUp( i );
                    continue;
                }
                
                //Exporter le block
                drawBlockShp( blk, myShapeFile, dbfHandle, field );
                
                //Progresser
                prog.moveUp( i );
                
                //Incrementer le nombre de block éxportés
                nb_el++;
            }
            
            //Liberer la mémoire
            DBFClose( dbfHandle );
            SHPClose( myShapeFile );
            
            // Messages
            print( "Exportation de " + to_string( nb_el ) + " block(s) dans " + acStrToStr( file ) + " terminée." );
            
        }
        
    }
    
    // Si aucun bloc n'est selectionné
    else
        print( "Aucun block sélectionné" );
        
    //Liberer la sélection
    acedSSFree( ssBlock );
    acedSSFree( ssCopie );
    
}

void cmdExportText()
{
    //Nom du texte
    ads_name ssText;
    
    //Fichier & Extension
    AcString file, ext;
    file = _T( "" );
    
    //Recuperer le  chemin du dwg courant
    AcString current_folder = getCurrentFileFolder();
    current_folder += _T( "\\" );
    
    //Tableau de stockage des données
    AcStringArray pData;
    
    //Message
    print( "Veuillez séléctionner le/les texte(s) : " );
    
    //Recuperer le nombre de texte
    long size = getSsText( ssText );
    
    // Si il y a des objets sélectionné
    if( size != 0 )
    {
        // Nombre d'objet
        int obj = 0;
        
        // Demander à l'utilisateur le repertoire du fichier à enregistrer
        file = askForFilePath(
                false,
                "xlsx;shp;csv;txt",
                "Enregistrer sous",
                current_folder
            );
            
        //Gestion d'erreur
        if( file == _T( "" ) )
        {
            print( "Commande annulée." );
            acedSSFree( ssText );
            return;
        }
        
        //Recuperer l'extension du fichier
        ext = getFileExt( file );
        
        //Changer l'encodage du chemin vers le fichier excel
        auto filename = latin1_to_utf8( acStrToStr( file ) );
        
        //L'entité texte multiple
        AcDbText* mTxt;
        
        // Si le fichier est de type .txt ou .csv
        if( ext == "txt" || ext == "csv" )
        {
            //Ouvrir le fichier
            fstream fichier( acStrToStr( file ), ios::app );
            
            //Ajouter les en-tête du fichier
            fichier << "Handle;Calque;Couleur;Type de ligne;Echelle type de ligne;Epaisseur de ligne;Transparence;Epaisseur;Materiau;X;Y;Z;Rotation;Valeur;Hauteur;Style;Mode verticale;Mode horizontale" << endl;
            
            //Barre de progression
            ProgressBar prog_txt = ProgressBar( _T( "Progression:" ), size );
            
            // Iteration sur les textes multiples sélectionnés
            for( long i = 0; i < size; i++ )
            {
                //Recuperer le i-eme MText
                mTxt = getTextFromSs( ssText, i, AcDb::kForRead );
                
                //verifier mtext
                if( !mTxt )
                {
                    prog_txt.moveUp( i );
                    continue;
                }
                
                //Recuperer les données
                exportText( mTxt, pData );
                
                // Iteration sur les données obtenus
                for( long j = 0; j < pData.size(); j++ )
                {
                    // Ecrire les données dans le tableau
                    fichier << pData[j] << ";";
                }
                
                // Se mettre à la ligne
                fichier << endl;
                
                // Vider le tableau des données
                pData.clear();
                
                // Incrémentation du nombre d'objet
                obj++;
                
                // Femeture de l'entité
                mTxt->close();
                
                //Progresser
                prog_txt.moveUp( i );
            }
            
            
            try
            {
                //Fermer le fichier
                fichier.close();
                
                // Affichage dans la console que l'exportation est terminé
                print( "Exportation de " + to_string( obj ) + " texte(s) dans " + acStrToStr( file ) + " terminée." );
                
                
            }
            
            catch( ... )
            {
                print( "Impossible d'enregistrer le fichier." );
                acedSSFree( ssText );
            }
        }
        
        //Si le fichier est de type .xlsx
        else if( ext == "xlsx" )
        {
            //Initialisation du fichier excel
            workbook wb;
            worksheet sheet1 = wb.active_sheet();
            
            //Ecrire les en-têtes du fichier excel
            sheet1.cell( 1, 1 ).value( "Handle" );
            sheet1.cell( 2, 1 ).value( "Calque" );
            sheet1.cell( 3, 1 ).value( "Couleur" );
            sheet1.cell( 4, 1 ).value( "Type de ligne" );
            sheet1.cell( 5, 1 ).value( "Echelle de type de ligne" );
            sheet1.cell( 6, 1 ).value( "Epaisseur de ligne" );
            sheet1.cell( 7, 1 ).value( "Transparence" );
            sheet1.cell( 8, 1 ).value( "Epaisseur" );
            sheet1.cell( 9, 1 ).value( "Materiau" );
            sheet1.cell( 10, 1 ).value( "X" );
            sheet1.cell( 11, 1 ).value( "Y" );
            sheet1.cell( 12, 1 ).value( "Z" );
            sheet1.cell( 13, 1 ).value( "Rotation" );
            sheet1.cell( 14, 1 ).value( "Valeur" );
            sheet1.cell( 15, 1 ).value( "Hauteur" );
            sheet1.cell( 16, 1 ).value( "Style" );
            sheet1.cell( 17, 1 ).value( "Mode verticale" );
            sheet1.cell( 18, 1 ).value( "Mode horizontale" );
            
            //Barre de progression
            ProgressBar prog = ProgressBar( _T( "Progression" ), size );
            
            //Iterer sur la selection
            for( long i = 0; i < size; i++ )
            {
                //Recuperer le i-eme AcDbMText
                mTxt = getTextFromSs( ssText, i, GcDb::kForRead );
                
                //Safeguard
                if( !mTxt )
                {
                    prog.moveUp( i );
                    continue;
                }
                
                //Recuperer les données
                exportText( mTxt, pData );
                
                //Mettre le format des cellules en texte
                sheet1.columns( true ).number_format( number_format::text() );
                
                //Iteration sur les données
                for( long k = 0; k < pData.size(); k++ )
                {
                    //Ecrire les données dans le fichier excel
                    
                    sheet1.cell( k + 1, i + 2 ).value( latin1_to_utf8( acStrToStr( pData[k] ) ) );
                }
                
                //Incrémenter le nombre d'objet
                obj++;
                
                //Reinitialiser le tableau
                pData.clear();
                
                //Liberer la mémoire
                mTxt->close();
                
                //Progresser
                prog.moveUp( i );
            }
            
            try
            {
                // Persister les données
                wb.save( filename );
                
                //Message
                print( "Exportation de " + to_string( obj ) + " texte(s) dans " + acStrToStr( file ) + " terminée." );
                
            }
            
            catch( ... )
            {
                print( "Impossible d'enregistrer le fichier." );
                acedSSFree( ssText );
            }
        }
        
        //Si le fichier est de type .shp
        else if( ext == "shp" )
        {
            //Lire le fichier de parametre
            vector<AcString> field = createTextField();
            writePrjOrNot( file );
            
            //Creer un fichier shp
            SHPHandle myShapeFile = SHPCreate( file, SHPT_POINTZ );
            
            //Creer un dbfFile
            DBFHandle dbfHandle = DBFCreate( file );
            
            //Creer les champs
            int fieldNumber = createFields( dbfHandle, field );
            
            //Barre de progression
            ProgressBar prog = ProgressBar( _T( "Progression :" ), size );
            
            //Nombre d'élements exportés
            int nb_el = 0;
            
            //Boucler sur toute les textes
            for( int i = 0; i < size; i++ )
            {
                //Recuperer l'objet block
                mTxt = getTextFromSs( ssText, i, AcDb::kForRead );
                
                //Verifier txt
                if( !mTxt )
                {
                    prog.moveUp( i );
                    continue;
                }
                
                //Exporter les textes
                drawTextShp( mTxt, myShapeFile, dbfHandle, field );
                
                //Progresser
                prog.moveUp( i );
                
                //Incrementer le nombre d'element
                nb_el++;
            }
            
            //Liberer la mémoire
            DBFClose( dbfHandle );
            SHPClose( myShapeFile );
            
            //Message
            print( "Exportation de " + to_string( nb_el ) + " texte(s) dans " + acStrToStr( file ) + " terminée." );
            
        }
    }
    
    //Si la sélection est vide
    else
        //Affichage dans la console
        print( "Aucun texte sélectionné." );
        
        
    //Liberer la selection
    acedSSFree( ssText );
}

void cmdExportMText()
{

    //Définir la selection
    ads_name ssMText;
    
    // Le fichier et l'extension
    AcString file, ext;
    
    file = _T( "" );
    
    //Recuperer le  chemin du dwg courant
    AcString current_folder = getCurrentFileFolder();
    current_folder += _T( "\\" );
    
    // Tableau de stockage des données
    AcStringArray pData;
    
    //Demander à l'utilisateur de sélectionner les textes
    print( "Veuillez sélectionner les textes multilignes : " );
    
    // Prendre le nombre d'objet sélectionnés
    long size = getSsMText( ssMText );
    
    // Si il y a des objets sélectionné
    if( size != 0 )
    {
        // Nombre d'objet
        int obj = 0;
        
        // Demander à l'utilisateur le repertoire du fichier à enregistrer
        file = askForFilePath(
                false,
                "xlsx;shp;csv;txt",
                "Enregistrer sous",
                current_folder
            );
            
        //Gestion d'erreur
        if( file == _T( "" ) )
        {
            print( "Commande annulée." );
            acedSSFree( ssMText );
            return;
        }
        
        //Recuperer l'extension du fichier
        ext = getFileExt( file );
        
        //Changer l'encodage du chemin vers le fichier excel
        auto filename = latin1_to_utf8( acStrToStr( file ) );
        
        //L'entité texte multiple
        AcDbMText* mTxt;
        
        // Si le fichier est de type .txt ou .csv
        if( ext == "txt" || ext == "csv" )
        {
            //Ouvrir le fichier
            fstream fichier( acStrToStr( file ), ios::app );
            
            //Ajouter les en-tête du fichier
            fichier << "Handle;Calque;Couleur;Type de ligne;Echelle de type de ligne;Epaisseur de ligne;Transparence;Taille du texte;Materiaux;Point d'insertion;Rotation;Contenu;Hauteur;Style;Facteur d'expace de ligne;Style d'espace de ligne;Flux de direction" << endl;
            
            
            //Barre de progression
            ProgressBar prog_txt = ProgressBar( _T( "Progression:" ), size );
            
            // Iteration sur les textes multiples sélectionnés
            for( long i = 0; i < size; i++ )
            {
                //Recuperer le i-eme MText
                mTxt = getMTextFromSs( ssMText, i, AcDb::kForRead );
                
                //verifier mTxt
                if( !mTxt )
                    prog_txt.moveUp( i );
                    
                //Recuperer les données
                exportMText( mTxt, pData );
                
                // Iteration sur les données obtenus
                for( long j = 0; j < pData.size(); j++ )
                {
                    // Ecrire les données dans le tableau
                    fichier << pData[j] << ";";
                }
                
                // Se mettre à la ligne
                fichier << endl;
                
                // Vider le tableau des données
                pData.clear();
                
                // Incrémentation du nombre d'objet
                obj++;
                
                // Femeture de l'entité
                mTxt->close();
                
                //Progresser
                prog_txt.moveUp( i );
            }
            
            
            try
            {
                //Fermer le fichier
                fichier.close();
                
                // Affichage dans la console que l'exportation est terminé
                print( "Exportation de " + to_string( obj ) + " texte(s) multiligne dans " + acStrToStr( file ) + " terminée." );
                
                // Réinitialisation du nombre d'objet
                obj = 0;
                
            }
            
            catch( ... )
            {
                print( "Impossible d'enregistrer le fichier." );
            }
        }
        
        //Si le fichier est de type .xlsx
        else if( ext == "xlsx" )
        {
            //Initialisation du fichier excel
            workbook wb;
            worksheet sheet1 = wb.active_sheet();
            
            //Ecrire les en-têtes du fichier excel
            sheet1.cell( 1, 1 ).value( "Handle" );
            sheet1.cell( 2, 1 ).value( "Calque" );
            sheet1.cell( 3, 1 ).value( "Couleur" );
            sheet1.cell( 4, 1 ).value( "Type de ligne" );
            sheet1.cell( 5, 1 ).value( "Echelle de type de ligne" );
            sheet1.cell( 6, 1 ).value( "Epaisseur de ligne" );
            sheet1.cell( 7, 1 ).value( "Transparence" );
            sheet1.cell( 8, 1 ).value( "Taille du texte" );
            sheet1.cell( 9, 1 ).value( "Materiaux" );
            sheet1.cell( 10, 1 ).value( "Point d'insertion" );
            sheet1.cell( 11, 1 ).value( "Rotation" );
            sheet1.cell( 12, 1 ).value( "Contenu" );
            sheet1.cell( 13, 1 ).value( "Hauteur" );
            sheet1.cell( 14, 1 ).value( "Style" );
            sheet1.cell( 15, 1 ).value( "Facteur d'espace de ligne" );
            sheet1.cell( 16, 1 ).value( "Style d'espace de ligne" );
            sheet1.cell( 17, 1 ).value( "Flux de direction" );
            
            //Barre de progression
            ProgressBar prog = ProgressBar( _T( "Progression" ), size );
            
            //Iterer sur la selection
            for( long i = 0; i < size; i++ )
            {
                // L'entité MText
                AcDbMText* mTxt;
                
                //Recuperer le i-eme AcDbMText
                mTxt = getMTextFromSs( ssMText, i, GcDb::kForRead );
                
                //Safeguard
                if( !mTxt )
                {
                    prog.moveUp( i );
                    continue;
                }
                
                //Recuperer les données
                exportMText( mTxt, pData );
                
                //Mettre le format des cellules en texte
                sheet1.columns( true ).number_format( number_format::text() );
                
                //Iteration sur les données
                for( long k = 0; k < pData.size(); k++ )
                {
                    //Ecrire les données dans le fichier excel
                    
                    sheet1.cell( k + 1, i + 2 ).value( latin1_to_utf8( acStrToStr( pData[k] ) ) );
                }
                
                //Incrémenter le nombre d'objet
                obj++;
                
                //Reinitialiser le tableau
                pData.clear();
                
                //Liberer la mémoire
                mTxt->close();
                
                //Progresser
                prog.moveUp( i );
            }
            
            try
            {
                // Persister les données
                wb.save( filename );
                
                //Message
                print( "Exportation de " + to_string( obj ) + " texte(s) multiligne dans " + acStrToStr( file ) + " terminée." );
                
            }
            
            catch( ... )
            {
                print( "Impossible d'enregistrer le fichier." );
                acedSSFree( ssMText );
            }
        }
        
        //Si le fichier est de type .shp
        else if( ext == "shp" )
        {
            //Lire le fichier de parametre
            vector<AcString> field = createTextField();
            writePrjOrNot( file );
            
            //Creer un fichier shp
            SHPHandle myShapeFile = SHPCreate( file, SHPT_POINTZ );
            
            //Creer un dbfFile
            DBFHandle dbfHandle = DBFCreate( file );
            
            //Creer les champs
            int fieldNumber = createFields( dbfHandle, field );
            
            // Initialistion de la barre de progression
            ProgressBar prog = ProgressBar( _T( "Progression" ), size );
            
            //Nombre d'elements exportés
            int nb_el = 0;
            
            //Boucler sur toute les textes
            for( int i = 0; i < size; i++ )
            {
                //Recuperer l'objet block
                mTxt = getMTextFromSs( ssMText, i, AcDb::kForRead );
                
                AcDbVoidPtrArray selectionSet;
                mTxt->explode( selectionSet );
                
                for( int k = 0; k < selectionSet.size(); k++ )
                {
                    AcDbText* text = ( AcDbText* )selectionSet[k];
                    
                    //Verifier txt
                    if( !text )
                        continue;
                        
                    drawTextShp( text, myShapeFile, dbfHandle, field );
                    
                    //Incrementer le nombre d'objet
                    nb_el++;
                }
                
                // incrementer la barre de progression
                mTxt->close();
                
                //Progresser
                prog.moveUp( i );
            }
            
            //Liberer la mémoire
            DBFClose( dbfHandle );
            SHPClose( myShapeFile );
            
            //Message
            print( "Exportation de " + to_string( nb_el ) + " texte(s) multiligne(s) dans " + acStrToStr( file ) + " terminée." );
            
        }
        
    }
    
    //Si la sélection est vide
    else
        //Affichage dans la console
        print( "Aucun texte multiligne sélectionné" );
        
    //Liberer la mémoire
    acedSSFree( ssMText );
}

void cmdExportEntity()
{
    //Définir la selection
    ads_name ssEnt;
    
    //Fichie & extenstion
    AcString file = _T( "" );
    AcString ext;
    
    //Recuperer le  chemin du dwg courant
    AcString current_folder = getCurrentFileFolder();
    current_folder += _T( "\\" );
    
    //Tableau de stockage des données
    AcStringArray pData;
    
    
    //Message
    print( "Selectionner les entités" );
    
    //Selectionner les entités
    long len_entity = getSsObject( ssEnt );
    
    if( len_entity > 0 )
    {
        //Nombre d'objet séléctionnés
        int obj = 0;
        
        //Demander le chemin du fichier d'enregistrement
        file = askForFilePath( false,
                "xlsx",
                "Enregistrer sous",
                current_folder );
                
        //Prendre l'extension du fichier
        ext = getFileExt( file );
        
        //Change file path encodage
        auto filename = latin1_to_utf8( acStrToStr( file ) );
        
        
        //L'entité polyligne 3d
        AcDb3dPolyline* pl3d;
        
        if( ext == "xlsx" )
        {
            //Initialisation du fichier excel
            workbook wb2;
            worksheet ws = wb2.active_sheet();
            
            //Ajout des en-tête du feuille
            ws.cell( 1, 1 ).value( "Handle" );
            ws.cell( 2, 1 ).value( "Calque" );
            ws.cell( 3, 1 ).value( "Couleur" );
            ws.cell( 4, 1 ).value( "Type de ligne" );
            ws.cell( 5, 1 ).value( "Echelle de type de ligne" );
            ws.cell( 6, 1 ).value( "Epaisseur de ligne" );
            ws.cell( 7, 1 ).value( "Transparence" );
            ws.cell( 8, 1 ).value( "Epaisseur" );
            ws.cell( 9, 1 ).value( "Materiau" );
            
            //L'entité polyligne 3d
            AcDbObjectId obj_id = NULL;
            AcDbEntity* ent = NULL;
            
            ProgressBar prog_xls = ProgressBar( _T( "Progression: " ), len_entity );
            
            for( long i = 0; i < len_entity; i++ )
            {
            
                //Le pointeur sur la polyligne
                obj_id = getObIdFromSs( ssEnt, i );
                
                //Ouvrir l'entité
                acdbOpenAcDbEntity( ent, obj_id, AcDb::kForRead );
                
                //Verifier ent
                if( !ent )
                    continue;
                    
                // Appel de la fonction de traitement
                exportEntity( ent, pData );
                
                // Mettre le format de toute les colonnes en texte
                ws.columns( true ).number_format( number_format::text() );
                
                // Iteration sur le tableau des données
                for( long k = 0; k < 10; k++ )
                {
                    //Ecriture des données dans le fichier excel
                    ws.cell( k + 1, i + 2 ).value( pData[k] );
                }
                
                //Incrémentation du nombre d'objet
                obj++;
                
                // Vider le tableau des données
                pData.clear();
                
                //Fermer la polyligne
                ent->close();
                prog_xls.moveUp( i );
            }
            
            try
            {
                //sauvegarder les modification
                wb2.save( filename );
                
                //Affichage dans la console que les informations sont bien exportés
                print( "Exportation de " + to_string( obj ) + " entités dans " + acStrToStr( file ) + " terminée." );
            }
            
            catch( ... )
            {
                print( "Impossible d'enregistrer les modifications." );
            }
        }
    }
    
    else
        print( "Aucune entité selectionnée." );
        
    //Liberer la selection
    acedSSFree( ssEnt );
}

void cmdExportPoly3DToXml()
{
    //Recuperer tous les noms des calques
    vector<AcString> layerNames = getLayerList();
    
    //Recuperer la taille des calques
    long laySize = layerNames.size();
    
    //Recuperer le  chemin du dwg courant
    AcString current_folder = getCurrentFileFolder();
    current_folder += _T( "\\" );
    
    //Verifier
    if( laySize == 0 )
    {
        //Afficher message
        print( "Aucun calque recuperé, Sortie" );
        return;
    }
    
    //2 - Demander à l'utilisateur le répertoire, le nom et l'extension du fichier ou il va enregistrer dans un dossier
    AcString xmlFilePath = askForFilePath( false,
            "xml",
            "Enregistrer sous (Nom de dossier sans espace)",
            current_folder
        );
        
    if( xmlFilePath == _T( "" ) )
    {
        //Afficher message
        print( "Commande annulée" );
        
        return;
    }
    
    //Recuperer le dossier du dwg
    AcString dwgFolder = getCurrentFileFolder();
    
    //Creer un ofstream
    ofstream xmlFile( acStrToStr( xmlFilePath ), std::ios::app );
    
    //Ecrire l'entete dans le fichier
    writeXmlHeader( xmlFile );
    
    //Selection sur les polylignes 3d
    ads_name ssPoly3d;
    
    //Barre de progression
    ProgressBar prog = ProgressBar( _T( "Progression" ), laySize );
    
    //Boucle sur les calques
    for( int c = 0; c < laySize; c++ )
    {
        //Recuperer le c-eme calque
        AcString cLayer = layerNames[c];
        
        //Recuperer les polylignes 3d dans le calque
        long lenPoly3d = getSsAllPoly3D( ssPoly3d, cLayer );
        
        //Verifier
        if( lenPoly3d == 0 )
            continue;
            
        //Polylignes 3d
        AcDb3dPolyline* poly3d = NULL;
        
        //Boucle sur les polylignes
        for( int p = 0; p < lenPoly3d; p++ )
        {
            //Recuperer le p-eme polyligne
            poly3d = getPoly3DFromSs( ssPoly3d, p );
            
            //Verifier
            if( !poly3d )
                continue;
                
            //Decomposer la polyligne en segment
            vector<LinePoly> segPoly = poly3dToSeg( poly3d );
            
            //Fermer la polyligne
            poly3d->close();
            
            //Ecrire dans le fichier XML
            exportPoly3dToXml( xmlFile,
                segPoly,
                cLayer,
                lenPoly3d,
                p );
        }
        
        //Progresser
        prog.moveUp( c );
    }
    
    //Ecrire le pied du xml
    writeXmlFooter( xmlFile );
    
    //Fermer le xml
    xmlFile.close();
    
    //Liberer la mémoire
    acedSSFree( ssPoly3d );
}

void cmdMoinsFaceToXml()
{
    //On ouvre le fichier
    AcString fileName = curDoc()->fileName();
    
    acutPrintf( fileName );
    
    AcString path, name, ext;
    
    splitPath( fileName, path, name, ext );
    
    ofstream file( acStrToStr( path ) + acStrToStr( name ) + ".xml" );
    
    //Ecrire l'entete
    file << "<?xml version=\"1.0\"\?>";
    file << "\n<LandXML xmlns=\"http://www.landxml.org/schema/LandXML-1.2\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:schemaLocation=\"http://www.landxml.org/schema/LandXML-1.2 http://www.landxml.org/schema/LandXML-1.2/LandXML-1.2.xsd\" date=\"2019-08-13\" time=\"11:16:50\" version=\"1.2\" language=\"English\" readOnly=\"false\">";
    file << "\n<Units><Metric areaUnit=\"squareMeter\" linearUnit=\"meter\" volumeUnit=\"cubicMeter\" temperatureUnit=\"celsius\" pressureUnit=\"milliBars\" diameterUnit=\"millimeter\" angularUnit=\"decimal degrees\" directionUnit=\"decimal degrees\"></Metric></Units>";
    file << "\n<Project name=\"MNT\"></Project>";
    file << "\n<Surfaces>";
    
    int iFace = 0;
    vector< AcString > layers = getLayerList();
    int size = layers.size();
    
    for( int i = 0; i < size; i++ )
        iFace += exportFaceToXml( &file, acStrToStr( layers[i] ) );
        
    //Terminer le fichier
    file << "\n</Surfaces>";
    file << "\n</LandXML>";
    
    // on ferme le fichier
    file.close();
    
    //Afficher le resultat
    acutPrintf( _T( "\n%d faces traitees" ), iFace );
}

void cmdFaceToXml()
{
    //Définir la selection
    ads_name faceSelection;
    
    //Selectionner les faces
    long length = getSelectionSet( faceSelection, "", "3DFACE" );
    
    //Recuperer le  chemin du dwg courant
    AcString current_folder = getCurrentFileFolder();
    current_folder += _T( "\\" );
    
    //Verifier
    if( length == 0 )
    {
        //Message
        print( "Selection vide" );
        acedSSFree( faceSelection );
        return;
    }
    
    //Demander de selectionner le fichier xml
    AcString xmlPath = askForFilePath( false, "xml", _T( "Selectionner le fichier xml:" ), current_folder );
    
    //Verifier
    if( xmlPath == _T( "" ) )
    {
        //Message
        print( "Commande annulée." );
        
        //Liberer la mémoire
        acedSSFree( faceSelection );
        
        return;
    }
    
    //Ouvrir le fichier
    ofstream file( acStrToStr( xmlPath ) );
    
    //Ecrire l'entete
    file << "<?xml version=\"1.0\"\?>";
    file << "\n<LandXML xmlns=\"http://www.landxml.org/schema/LandXML-1.2\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:schemaLocation=\"http://www.landxml.org/schema/LandXML-1.2 http://www.landxml.org/schema/LandXML-1.2/LandXML-1.2.xsd\" date=\"2019-08-13\" time=\"11:16:50\" version=\"1.2\" language=\"English\" readOnly=\"false\">";
    file << "\n<Units><Metric areaUnit=\"squareMeter\" linearUnit=\"meter\" volumeUnit=\"cubicMeter\" temperatureUnit=\"celsius\" pressureUnit=\"milliBars\" diameterUnit=\"millimeter\" angularUnit=\"decimal degrees\" directionUnit=\"decimal degrees\"></Metric></Units>";
    file << "\n<Project name=\"MNT\"></Project>";
    file << "\n<Surfaces>";
    
    //Face
    AcDbFace* face = NULL;
    
    //Définir une selection pour l'entité
    ads_name ent;
    
    //Declarer un AcDbObjectId
    AcDbObjectId id = AcDbObjectId::kNull;
    
    //Recuperer tous les calques du dessin
    vector<AcString> layers = getLayerList();
    
    //Recuperer la taille des calques
    int laySize = layers.size();
    
    if( laySize == 0 )
    {
        //Message
        print( "Aucun calque dans le dessin" );
        
        //Liberer la selection
        acedSSFree( faceSelection );
        
        //Terminer la selection
        return;
    }
    
    //Vecteur de FacesInLayer
    vector<FacesInLayer> vecFL;
    
    //Iterer dans les calques
    for( int i = 0; i < laySize; i++ )
    {
        //Creer un faceLayer
        FacesInLayer tFL;
        tFL.layer = layers[i];
        
        //Ajouter le faceLayer dans le vecteur
        vecFL.push_back( tFL );
    }
    
    //Recuperer la taille du vecteur de facelayer
    int fSize = vecFL.size();
    
    //Barre de progression
    ProgressBar prog = ProgressBar( _T( "Preparation des faces" ), length );
    
    //Boucle sur les selections des faces
    for( int i = 0; i < length; i++ )
    {
        //Recuperer le i eme face
        if( acedSSName( faceSelection, i, ent ) != RTNORM )
            continue;
            
        if( acdbGetObjectId( id, ent ) != Acad::eOk )
            continue;
            
        AcDbEntity* pEnt = NULL;
        
        if( acdbOpenAcDbEntity( pEnt, id, AcDb::kForWrite ) != Acad::eOk )
            continue;
            
        AcDbFace* castEntity = static_cast<AcDbFace*>( pEnt );
        
        //Recuperer le calque de la face
        AcString layFace = castEntity->layer();
        
        //Boucle sur le vecteur de facelayer
        for( int f = 0; f < fSize; f++ )
        {
            //Verifier le calque
            if( layFace == vecFL[f].layer )
            {
                //Ajouter l'objectId de la face dans le vecteur
                vecFL[f].facesIds.push_back( castEntity->objectId() );
            }
        }
        
        //Fermer la face
        castEntity->close();
        
        //Progresser
        prog.moveUp( i );
    }
    
    //Appeler la fonction qui va exporter les faces dans le xml
    exportFaceToXml( &file, vecFL );
    
    //Terminer le fichier
    file << "\n</Surfaces>";
    file << "\n</LandXML>";
    
    //On ferme le fichier
    file.close();
    
    //Afficher le resultat
    acutPrintf( _T( "\n%d faces traitées." ), length );
    
    acedSSFree( faceSelection );
}

void cmdExportPoly()
{
    //Définir la sélection
    ads_name ssPoly;
    
    // Le fichier
    AcString file,
             // L'extension du fichier
             ext;
             
    file = _T( "" );
    
    //Recuperer le  chemin du dwg courant
    AcString current_folder = getCurrentFileFolder();
    current_folder += _T( "\\" );
    
    
    // Tableau de stockage des informations de l'entité
    AcStringArray pData;
    
    // Demander à l'utilisateur de séléctionner les polylignes
    print( "Veuillez sélectionner les polylignes : " );
    
    
    // Les sélection sur les polylignes
    long size = getSelectionSet( ssPoly, "", "LWPOLYLINE,POLYLINE" );
    
    // Verification si la sélection n'est pas vide
    if( size != 0 )
    {
        // Nombre d'objet sélectionnées
        int obj = 0;
        
        // Demander à l'utilisateur le repertoire et l'extension du fichier à enregistrer
        file = askForFilePath( false,
                "xlsx;shp;csv;txt",
                "Enregistrer sous",
                current_folder );
                
        //Gestion d'erreur
        if( file == _T( "" ) )
        {
            print( "Commande annulée." );
            acedSSFree( ssPoly );
            return;
        }
        
        // Prendre l'extension fu fichier
        ext = getFileExt( file );
        
        //Changer l'encodage du chemin pour le fichier excel
        auto filename = latin1_to_utf8( acStrToStr( file ) );
        
        // L'entité polyligne 2d
        AcDbObjectId obj_id;
        AcDbEntity* ent = NULL;
        
        // Si le fichier est .txt ou .csv
        if( ext == "txt" || ext == "csv" )
        {
            // Flux de fichier
            fstream fichier( acStrToStr( file ), ios::app );
            
            // Ecriture des en-têtes du fichier
            fichier << "Handle;Calque;Couleur;Type de ligne;Echelle de type de ligne;Epaisseur de ligne;Transparence;Epaisseur;Materiau;Longueur 2d;Largeur;Fermée" << endl;
            
            ProgressBar prog_txt = ProgressBar( _T( "Progression:" ), size );
            
            // Iteration sur les objets
            for( long i = 0; i < size; i++ )
            {
                // Le pointeur sur la polyligne
                obj_id = getObIdFromSs( ssPoly, i );
                
                //Ouvrir l'entité
                acdbOpenAcDbEntity( ent, obj_id, AcDb::kForRead );
                
                //verifier ent
                if( !ent )
                {
                    prog_txt.moveUp( i );
                    continue;
                }
                
                // Appel de la fonction de traitement
                exportPoly( ent, pData );
                
                // Iteration sur le tableau des données
                for( long j = 0; j < pData.size(); j++ )
                {
                    // Ecriture des données dans le fichier
                    fichier << pData[j] << ";";
                }
                
                // Se mettre à la ligne
                fichier << endl;
                
                // Incrémentation du nomber d'objet
                obj++;
                
                // Vider le tableau des données
                pData.clear();
                
                // Fermer la polyligne
                ent->close();
                prog_txt.moveUp( i );
            }
            
            try
            {
                // Fermer le flux de fichier
                fichier.close();
                
                // Affichage dans la console que les informations sont bien exportés
                print( "Exportation de " + to_string( obj ) + " polyligne(s) 2D dans " + acStrToStr( file ) + " terminée." );
                
                // Réinitialisation du nombre d'objet
                obj = 0;
            }
            
            catch( ... )
            {
                print( "Impossible d'enregistrer le fichier." );
            }
        }
        
        // Si le fichier est .xlsx(excel)
        else if( ext == "xlsx" )
        {
            // Initialisation du fichier excel
            workbook wb;
            worksheet ws = wb.active_sheet();
            
            // Ajout des en-têtes dans le feuille excel
            ws.cell( 1, 1 ).value( "Handle" );
            ws.cell( 2, 1 ).value( "Calque" );
            ws.cell( 3, 1 ).value( "Couleur" );
            ws.cell( 4, 1 ).value( "Type de ligne" );
            ws.cell( 5, 1 ).value( "Echelle de type de ligne" );
            ws.cell( 6, 1 ).value( "Epaisseur de ligne" );
            ws.cell( 7, 1 ).value( "Transparence" );
            ws.cell( 8, 1 ).value( "Epaisseur" );
            ws.cell( 9, 1 ).value( "Materiau" );
            ws.cell( 10, 1 ).value( "Longueur 2d" );
            ws.cell( 11, 1 ).value( "Longueur 3d" );
            ws.cell( 12, 1 ).value( u8"Fermée" );
            
            
            ProgressBar prog_xls = ProgressBar( _T( "Progression:" ), size );
            
            //incrementeur de colonne
            int col_nb = 0;
            
            // Iteration sur les objets
            for( long i = 0; i < size; i++ )
            {
                //Recuperer le i-eme id de l'entité
                obj_id = getObIdFromSs( ssPoly, i );
                
                //Ouvrir l'entité
                acdbOpenAcDbEntity( ent, obj_id, AcDb::kForRead );
                
                //verifier ent
                if( !ent )
                {
                    prog_xls.moveUp( i );
                    continue;
                }
                
                // Appel de la fonction de traitement
                exportPoly( ent, pData );
                
                
                // Mettre le format de toute les cellules en Texte
                ws.columns( true ).number_format( number_format::text() );
                
                // Iteration sur le tableau des données
                for( long j = 0; j < pData.size(); j++ )
                {
                    //Changer l'encodage des données en utf-8
                    auto data = latin1_to_utf8( acStrToStr( pData[j] ) );
                    
                    //Ecriture des données dans le fichier
                    ws.cell( j + 1, col_nb + 2 ).value( data );
                }
                
                // Incrémentation du nombre d'objet
                obj++;
                col_nb++;
                
                // Vider le tableau des données
                pData.clear();
                
                // Fermer la polyligne 2d
                ent->close();
                prog_xls.moveUp( i );
            }
            
            try
            {
                // Sauvegarder le fichier excel .xlsx
                wb.save( filename );
                
                // Affichage dans la console que les propriétes sont bien exportés
                print( "Exportation de " + to_string( obj ) + " polyligne(s) dans " + acStrToStr( file ) + " terminée." );
                
                // Réinitialisation de obj
                obj = 0;
            }
            
            catch( ... )
            {
                print( "Impossible d'enregistrer le fichier." );
                
                //liberer la mémoire
                acedSSFree( ssPoly );
            }
            
        }
        
        //Si le fichier est .shp (shape)
        else if( ext == "shp" )
        {
            //Lire le fichier de parametre
            vector<AcString> field = createPolyField();
            writePrjOrNot( file );
            
            //Creer un fichier shp
            SHPHandle myShapeFile = SHPCreate( file, SHPT_ARCZ );
            
            //Creer un dbfFile
            DBFHandle dbfHandle = DBFCreate( file );
            
            //Creer les champs
            int fieldNumber = createFields( dbfHandle, field );
            
            
            // Initialistion de la barre de progression
            ProgressBar prog = ProgressBar( _T( "Progression" ), size );
            
            //Nombre d'objets exportés
            int nb_el = 0;
            
            //Boucler sur la selection
            for( int i = 0; i < size; i++ )
            {
                //Recuperer le i-eme id de l'entité
                obj_id = getObIdFromSs( ssPoly, i );
                
                //Ouvrir l'entité
                acdbOpenAcDbEntity( ent, obj_id, AcDb::kForRead );
                
                //verifier ent
                if( !ent )
                {
                    prog.moveUp( i );
                    continue;
                }
                
                if( ent->isKindOf( AcDbPolyline::desc() ) )
                {
                    //Caster l'entité
                    AcDbPolyline* pl = AcDbPolyline::cast( ent );
                    
                    //Reciperer l'objet ligne
                    drawLineShp( pl, myShapeFile, dbfHandle, field );
                }
                
                else if( ent->isKindOf( AcDb3dPolyline::desc() ) )
                {
                    //Caster l'entité
                    AcDb3dPolyline* pl = AcDb3dPolyline::cast( ent );
                    
                    //Reciperer l'objet ligne
                    drawLineShp( pl, myShapeFile, dbfHandle, field );
                }
                
                //Fermer la polyline
                ent->close();
                
                // incrementer la barre de progression
                prog.moveUp( i );
                
                //Incrementer le nombre d'objet éxportés
                nb_el++;
            }
            
            //Liberer la mémoire
            DBFClose( dbfHandle );
            SHPClose( myShapeFile );
            
            //Affichage dans la console que les propriétes sont bien exportés
            print( "Exportation de " + to_string( nb_el ) + " polyligne(s) 2D dans " + acStrToStr( file ) + " terminée." );
        }
        
    }
    
    // Si la sélection est vide
    else
    {
        // Affichage dans la console
        print( "Aucune polyligne sélectionnée" );
    }
    
    
    acedSSFree( ssPoly );
}

void cmdExportClosedPoly2d()
{
    // Nom de la polyligne 2d
    ads_name ssPoly2d;
    
    // Le fichier
    AcString file,
             // L'extension du fichier
             ext;
             
    file = _T( "" );
    
    //Recuperer le  chemin du dwg courant
    AcString current_folder = getCurrentFileFolder();
    current_folder += _T( "\\" );
    
    // Tableau de stockage des données des entités
    AcStringArray pData;
    
    // Demander à l'utilisateur de séléctionner les polylignes
    print( "Veuillez sélectionner les polylignes : " );
    
    // Les sélection sur les polylignes
    long size = getSsPoly2D( ssPoly2d, "", true );
    
    // Verification si la sélection n'est pas vide
    if( size != 0 )
    {
        // Nombre d'objet sélectionnées
        int obj = 0;
        
        // Demander à l'utilisateur le repertoire et l'extension du fichier à enregistrer
        file = askForFilePath( false,
                "xlsx;shp;csv;txt",
                "Enregistrer sous",
                current_folder );
                
        //Gestion d'erreur
        if( file == _T( "" ) )
        {
            acedSSFree( ssPoly2d );
            print( "Commande annulée." );
            return;
        }
        
        // Prendre l'extension fu fichier
        ext = getFileExt( file );
        
        //Changer l'encodage du chemin pour le fichier excel
        auto filename = latin1_to_utf8( acStrToStr( file ) );
        
        // L'entité polyligne 2d
        AcDbPolyline* pl2d;
        
        // Si le fichier est .txt ou .csv
        if( ext == "txt" || ext == "csv" )
        {
            // Flux de fichier
            fstream fichier( acStrToStr( file ), ios::app );
            
            // Ecriture des en-têtes du fichier
            fichier << "Handle;Calque;Couleur;Type de ligne;Echelle de type de ligne;Epaisseur de ligne;Transparence;Epaisseur;Materiau;Longueur 2d;Surface" << endl;
            
            ProgressBar prog_txt = ProgressBar( _T( "Progression:" ), size );
            
            // Iteration sur les objets
            for( long i = 0; i < size; i++ )
            {
                // Le pointeur sur la polyligne
                pl2d = getPoly2DFromSs( ssPoly2d, i, AcDb::kForRead );
                
                //Verifier pl2d
                if( !pl2d )
                    continue;
                    
                // Appel de la fonction de traitement
                exportPoly2d( pl2d, pData );
                
                //Exclure la largeur
                pData.removeAt( 10 );
                
                //Exclure le statut
                pData.removeAt( 10 );
                
                // Iteration sur le tableau des données
                for( long j = 0; j < pData.size(); j++ )
                {
                    // Ecriture des données dans le fichier
                    fichier << pData[j] << ";";
                }
                
                // Se mettre à la ligne
                fichier << endl;
                
                // Incrémentation du nomber d'objet
                obj++;
                
                // Vider le tableau des données
                pData.clear();
                
                // Fermer la polyligne
                pl2d->close();
                
                //Progresser
                prog_txt.moveUp( i );
            }
            
            try
            {
                // Fermer le flux de fichier
                fichier.close();
                
                // Affichage dans la console que les informations sont bien exportés
                print( "Exportation de " + to_string( obj ) + " polyligne(s) 2D dans " + acStrToStr( file ) + " terminée." );
                
                // Réinitialisation du nombre d'objet
                obj = 0;
            }
            
            catch( ... )
            {
                acedSSFree( ssPoly2d );
                print( "Impossible d'enregistrer le fichier." );
                return;
            }
        }
        
        // Si le fichier est .xlsx(excel)
        else if( ext == "xlsx" )
        {
            // Initialisation du fichier excel
            workbook wb;
            worksheet ws = wb.active_sheet();
            
            // Ajout des en-têtes dans le feuille excel
            ws.cell( 1, 1 ).value( "Handle" );
            ws.cell( 2, 1 ).value( "Calque" );
            ws.cell( 3, 1 ).value( "Couleur" );
            ws.cell( 4, 1 ).value( "Type de ligne" );
            ws.cell( 5, 1 ).value( "Echelle de type de ligne" );
            ws.cell( 6, 1 ).value( "Epaisseur de ligne" );
            ws.cell( 7, 1 ).value( "Transparence" );
            ws.cell( 8, 1 ).value( "Epaisseur" );
            ws.cell( 9, 1 ).value( "Materiau" );
            ws.cell( 10, 1 ).value( "Longueur 2d" );
            ws.cell( 11, 1 ).value( "Surface" );
            
            
            // L'entité polyligne 2d
            AcDbPolyline* pl2d;
            
            ProgressBar prog_xls = ProgressBar( _T( "Progression:" ), size );
            
            // Iteration sur les objets
            for( long i = 0; i < size; i++ )
            {
                // Le pointeur sur la polyligne
                pl2d = getPoly2DFromSs( ssPoly2d, i, AcDb::kForRead );
                
                // Safeguard
                if( !pl2d )
                    continue;
                    
                // Appel de la fonction de traitement
                exportPoly2d( pl2d, pData );
                
                //Exclure la largeur
                pData.removeAt( 10 );
                
                //Exclure le statut
                pData.removeAt( 10 );
                
                // Mettre le format de toute les cellules en Texte
                ws.columns( true ).number_format( number_format::text() );
                
                // Iteration sur le tableau des données
                for( long j = 0; j < pData.size(); j++ )
                {
                    //Convertir l'encodage des strings
                    auto data = latin1_to_utf8( acStrToStr( pData[j] ) );
                    
                    // Ecriture des données dans le fichier
                    ws.cell( j + 1, i + 2 ).value( data );
                }
                
                // Incrémentation du nombre d'objet
                obj++;
                
                // Vider le tableau des données
                pData.clear();
                
                // Fermer la polyligne 2d
                pl2d->close();
                prog_xls.moveUp( i );
            }
            
            try
            {
                // Sauvegarder le fichier excel .xlsx
                wb.save( filename );
                
                // Affichage dans la console que les propriétes sont bien exportés
                print( "Exportation de " + to_string( obj ) + " polyligne(s) 2D dans " + acStrToStr( file ) + " terminée." );
                
                // Réinitialisation de obj
                obj = 0;
            }
            
            catch( ... )
            {
                acedSSFree( ssPoly2d );
                print( "Impossible d'enregistrer le fichier." );
                return;
            }
            
        }
        
        //Si le fichier est .shp(shape)
        
        else if( ext == "shp" )
        {
        
        
            //Lire le fichier de parametre
            vector<AcString> field = createClosedPolyField();
            writePrjOrNot( file );
            
            //Creer un fichier shp
            SHPHandle myShapeFile = SHPCreate( file, SHPT_POLYGONZ );
            
            //Creer un dbfFile
            DBFHandle dbfHandle = DBFCreate( file );
            
            //Creer les champs
            int fieldNumber = createFields( dbfHandle, field );
            
            
            // Initialistion de la barre de progression
            ProgressBar prog = ProgressBar( _T( "Progression" ), size );
            
            //Boucler sur la selection
            for( int i = 0; i < size; i++ )
            {
            
                //Recuperer l'entite
                AcDbPolyline* poly2d = getPoly2DFromSs( ssPoly2d, i );
                
                //Verifier ent
                if( !poly2d )
                    continue;
                    
                //recuperer le nombre de sommets
                int verTnumber = poly2d->numVerts();
                
                vector<AcGePoint3d> point;
                AcGePoint3d p;
                
                //Boucler sur les vecteurs
                for( int k = 0; k < verTnumber; k++ )
                {
                    poly2d->getPointAt( k, p );
                    point.push_back( p );
                }
                
                //Recuperer le point sur le sommet numero
                poly2d->getPointAt( 0, p );
                point.push_back( p );
                
                //Reciperer l'objet ligne
                drawClosedPolyShp( poly2d, point, myShapeFile, dbfHandle, field );
                
                poly2d->close();
                
                // incrementer la barre de progression
                
                prog.moveUp( i );
            }
            
            //Fermer la selection
            DBFClose( dbfHandle );
            SHPClose( myShapeFile );
            
            // Affichage dans la console que les propriétes sont bien exportés
            print( "Exportation de " + to_string( size ) + " polyligne(s) 2D dans " + acStrToStr( file ) + " terminée." );
            
        }
        
        
    }
    
    // Si la sélection est vide
    else
    {
        // Affichage dans la console
        print( "Aucune polyligne sélectionnée" );
    }
    
    // Liberer la sélection
    acedSSFree( ssPoly2d );
    
    
}

void cmdExportClosedPoly3d()
{
    //Nom de la polyligne 3d
    ads_name ssPoly3d;
    
    // Le fichier
    AcString file,
             // L'extension du fichier
             ext;
             
    file = _T( "" );
    
    //Recuperer le  chemin du dwg courant
    AcString current_folder = getCurrentFileFolder();
    current_folder += _T( "\\" );
    
    // Tableau de stockage des informations de l'entité
    AcStringArray pData;
    
    //1 - Demander à l'utilisateur à séléctionner la/les polyligne
    //Affichage dans la console en demandant à l'utilisateur de sélectionner les polyligne
    print( "Veuillez séléctionner la/les polyligne(s) 3D : " );
    
    long size = getSsPoly3D( ssPoly3d, "", true );
    
    //Vérification si il existe des polylignes séléctionnées
    if( size != 0 )
    {
    
        //Nombre d'objet séléctionnés
        int obj = 0;
        
        //2 - Demander à l'utilisateur le répertoire, le nom et l'extension du fichier ou il va enregistrer dans un dossier
        file = askForFilePath( false,
                "xlsx;shp;csv;txt",
                "Enregistrer sous",
                current_folder );
                
        //Gestion d'erreur
        if( file == _T( "" ) )
        {
            acedSSFree( ssPoly3d );
            print( "Commande annulée." );
            return;
        }
        
        //Prendre l'extension du fichier
        ext = getFileExt( file );
        
        //encoder en utf8
        auto filename = latin1_to_utf8( acStrToStr( file ) );
        
        //std::replace( filename.begin(), filename.end(), '\\', '/' ); // replace all '\' to '/'
        //L'entité polyligne 3d
        AcDb3dPolyline* pl3d;
        
        //Si c'est un fichier .csv ou un fichier .txt
        if( ext == "csv" || ext == "txt" )
        {
            //Ouvverture du fichier obtenue dans askForFilePath()
            fstream fichier( acStrToStr( file ), ios::app );
            
            //Ecriture des en-têtes du fichier
            fichier << "Handle;Calque;Couleur;Type de calque;Echelle de type de ligne;Epaisseur de ligne;Transparence;Materiau;Longueur 2d;Longueur 3d;Surface" << endl;
            
            //3 - Récupérer les informations sur la/les polyligne(s)
            
            // Iteration sur les objets
            for( long i = 0; i < size; i++ )
            {
            
                //Le pointeur sur la polyligne
                pl3d = getPoly3DFromSs( ssPoly3d, i, GcDb::kForRead );
                
                // Appel de la fonction de traitement
                exportPoly3d( pl3d, pData );
                
                // Iteration sur le tableau de données
                for( long j = 0; j < pData.size(); j++ )
                {
                    if( j == 10 )
                        continue;
                        
                    //Ecriture des données dans le fichier
                    fichier << pData[j] << ";";
                }
                
                // Se mettre à la ligne
                fichier << endl;
                
                //Incrémentation du nombre d'objet
                obj++;
                
                // Vider le tableau des données
                pData.clear();
                
                //Fermeture de la polyligne
                pl3d->close();
            }
            
            try
            {
                //Fermeture du fichier
                fichier.close();
                
                //Affichage dans le console que les informations sont bien exportés
                print( "Exportation de " + to_string( obj ) + " polylignes 3D dans " + filename + " terminée." );
                
                //Remettre la valeur de obj en 0
                obj = 0;
            }
            
            catch( ... )
            {
                print( "Impossible d'enregistrer le fichier." );
            }
        }
        
        //Si c'est un fichier Excel
        else if( ext == "xlsx" )
        {
            //Initialisation du fichier excel
            workbook wb2;
            worksheet ws = wb2.active_sheet();
            
            //Ajout des en-tête du feuille
            ws.cell( 1, 1 ).value( "Handle" );
            ws.cell( 2, 1 ).value( "Calque" );
            ws.cell( 3, 1 ).value( "Couleur" );
            ws.cell( 4, 1 ).value( "Type de ligne" );
            ws.cell( 5, 1 ).value( "Echelle de type de ligne" );
            ws.cell( 6, 1 ).value( "Epaisseur de ligne" );
            ws.cell( 7, 1 ).value( "Transparence" );
            ws.cell( 8, 1 ).value( "Materiau" );
            ws.cell( 9, 1 ).value( "Longueur 2d" );
            ws.cell( 10, 1 ).value( "Longueur 3d" );
            ws.cell( 11, 1 ).value( "Surface" );
            
            //L'entité polyligne 3d
            AcDb3dPolyline* pl3d;
            
            for( long i = 0; i < size; i++ )
            {
            
                //Le pointeur sur la polyligne
                pl3d = getPoly3DFromSs( ssPoly3d, i, GcDb::kForRead );
                
                //Safeguard
                if( !pl3d )
                    continue;
                    
                // Appel de la fonction de traitement
                exportPoly3d( pl3d, pData );
                
                // Mettre le format de toute les colonnes en texte
                ws.columns( true ).number_format( number_format::text() );
                
                // Iteration sur le tableau des données
                for( long k = 0; k < pData.size(); k++ )
                {
                    //Convertir l'encodage des strings
                    auto data = latin1_to_utf8( acStrToStr( pData[k] ) );
                    
                    if( k < 10 )
                        //Ecrire les données
                        ws.cell( k + 1, i + 2 ).value( data );
                        
                    //Exclure du fichier le status de la polyline
                    if( k == 10 )
                        continue;
                        
                    if( k > 10 )
                        ws.cell( k, i + 2 ).value( data );
                }
                
                
                
                //Incrémentation du nombre d'objet
                obj++;
                
                // Vider le tableau des données
                pData.clear();
                
                //Fermer la polyligne
                pl3d->close();
            }
            
            
            
            try
            {
                //Fermer le fichier
                wb2.save( filename );
                
                //Affichage dans la console que les informations sont bien exportés
                print( "Exportation de " + to_string( obj ) + " polyligne 3D dans " + filename + " terminée." );
                
                //Remettre la valeur de obj en 0
                obj = 0;
            }
            
            catch( ... )
            {
                print( "Impossible d'enregistrer le fichier." );
            }
            
        }
        
        //Si le fichier est .shp (shape)
        else if( ext == "shp" )
        {
            //Lire le fichier de parametre
            vector<AcString> field = createClosedPoly3dField();
            writePrjOrNot( file );
            
            //Creer un fichier shp
            SHPHandle myShapeFile = SHPCreate( file, SHPT_POLYGONZ );
            
            //Creer un dbfFile
            DBFHandle dbfHandle = DBFCreate( file );
            
            //Creer les champs
            int fieldNumber = createFields( dbfHandle, field );
            
            
            // Initialistion de la barre de progression
            ProgressBar prog = ProgressBar( _T( "Progression" ), size );
            
            //Nombre d'objet exportés
            int nb_el = 0;
            
            //Boucler sur la selection
            for( int i = 0; i < size; i++ )
            {
            
                //Recuperer l'entite
                AcDb3dPolyline* poly = getPoly3DFromSs( ssPoly3d, i );
                
                //Verifier la polyline
                if( !poly )
                    continue;
                    
                vector<AcGePoint3d> point;
                
                AcGePoint3d p;
                //Boucer sur les sommets
                AcDbObjectIterator* iter = poly->vertexIterator();
                AcDb3dPolylineVertex* vertex;
                
                for( iter->start(); !iter->done(); iter->step() )
                {
                    if( !poly->openVertex( vertex, iter->objectId(), AcDb::kForRead ) )
                    {
                        p = vertex->position();
                        point.push_back( p );
                    }
                }
                
                vertex->close();
                //Recuperer le point sur le sommet numero 0
                poly->getStartPoint( p );
                point.push_back( p );
                
                //Reciperer l'objet ligne
                drawClosedPolyShp( poly, point, myShapeFile, dbfHandle, field );
                
                //Fermer la polyline
                poly->close();
                
                // incrementer la barre de progression
                prog.moveUp( i );
                
                //Incrementer le nombre d'objet exportés
                nb_el++;
            }
            
            //Liberer ma mémoire
            DBFClose( dbfHandle );
            SHPClose( myShapeFile );
            
            //Affichage dans la console que les informations sont bien exportés
            print( "Exportation de " + to_string( nb_el ) + " polyligne 3D dans " + acStrToStr( file ) + " terminée." );
            
        }
        
        
    }
    
    //Si il n'y a pas de polyline séléctionnée
    else
        print( "Aucune polyligne séléctionnée ! " );
        
    //Liberer la selection
    acedSSFree( ssPoly3d );
}

void cmdExportClosedPoly()
{
    //Définir la sélection
    ads_name ssPoly;
    
    //Variable qui contient le choix de l'utilisateur
    AcString u_choice = _T( "" );
    
    // Le fichier
    AcString file,
             // L'extension du fichier
             ext;
             
    file = _T( "" );
    
    //Recuperer le  chemin du dwg courant
    AcString current_folder = getCurrentFileFolder();
    current_folder += _T( "\\" );
    
    // Tableau de stockage des informations de l'entité
    AcStringArray pData;
    
    
    // Demander à l'utilisateur de séléctionner les polylignes
    print( "Veuillez sélectionner les polylignes : " );
    
    // Les sélection sur les polylignes
    long size = getSsObject( ssPoly, _T( "" ) );
    
    // Verification si la sélection n'est pas vide
    if( size != 0 )
    {
        // Nombre d'objet sélectionnées
        int obj = 0;
        
        // Demander à l'utilisateur le repertoire et l'extension du fichier à enregistrer
        file = askForFilePath( false,
                "xlsx;shp;csv;txt",
                "Enregistrer sous",
                current_folder );
                
        //Gestion d'erreur
        if( file == _T( "" ) )
        {
            acedSSFree( ssPoly );
            print( "Commande annulée." );
            return;
        }
        
        // Prendre l'extension fu fichier
        ext = getFileExt( file );
        
        //Changer l'encodage du chemin pour le fichier excel
        auto filename = latin1_to_utf8( acStrToStr( file ) );
        
        
        
        //AcDbObjectid
        AcDbObjectId obj_id;
        
        // Si le fichier est .txt ou .csv
        if( ext == "txt" || ext == "csv" )
        {
        
            // Flux de fichier
            fstream fichier( acStrToStr( file ), ios::app );
            
            // Ecriture des en-têtes du fichier
            fichier << "Handle;Calque;Couleur;Type de ligne;Echelle de type de ligne;Epaisseur de ligne;Transparence;Epaisseur;Materiau;Longueur 2d;Longueur 3d;Surface" << endl;
            
            ProgressBar prog_txt = ProgressBar( _T( "Progression:" ), size );
            
            // Iteration sur les objets
            for( long i = 0; i < size; i++ )
            {
                // Les entités polylines
                AcDbPolyline* pl2d = NULL;
                AcDb3dPolyline* pl3d = NULL;
                AcDbEntity* ent = NULL;
                
                
                // Le pointeur sur la polyligne
                obj_id = getObIdFromSs( ssPoly, i );
                
                //Ouvrir l'entité
                acdbOpenAcDbEntity( ent, obj_id, AcDb::kForRead );
                
                //verifier ent
                if( !ent )
                    continue;
                    
                if( ent->isKindOf( AcDbPolyline::desc() ) )
                {
                
                    //Recuperer la polyline
                    pl2d = AcDbPolyline::cast( ent );
                    
                    //veririfier pl2d
                    if( !pl2d )
                    {
                        prog_txt.moveUp( i );
                        ent->close();
                        continue;
                    }
                    
                    if( !isClosed( pl2d ) )
                    {
                    
                        prog_txt.moveUp( i );
                        pl2d->close();
                        continue;
                    }
                    
                    //Exporter l'entité
                    exportClosedPoly( ent, pData );
                }
                
                else if( ent->isKindOf( AcDb3dPolyline::desc() ) )
                {
                    //Recuperer la polyline
                    pl3d = AcDb3dPolyline::cast( ent );
                    
                    //veririfier pl2d
                    if( !pl3d )
                    {
                        prog_txt.moveUp( i );
                        ent->close();
                        continue;
                    }
                    
                    if( !isClosed( pl3d ) )
                    {
                        prog_txt.moveUp( i );
                        pl3d->close();
                        continue;
                    }
                    
                    //Exporter l'entité
                    exportClosedPoly( ent, pData );
                }
                
                else
                {
                    prog_txt.moveUp( i );
                    ent->close();
                    continue;
                }
                
                
                // Iteration sur le tableau des données
                for( long j = 0; j < pData.size(); j++ )
                {
                    // Ecriture des données dans le fichier
                    fichier << pData[j] << ";";
                }
                
                
                // Se mettre à la ligne
                fichier << endl;
                
                // Incrémentation du nomber d'objet
                obj++;
                
                // Vider le tableau des données
                pData.clear();
                
                
                prog_txt.moveUp( i );
            }
            
            try
            {
                // Fermer le flux de fichier
                fichier.close();
                
                // Affichage dans la console que les informations sont bien exportés
                print( "Exportation de " + to_string( obj ) + " polyligne(s) dans " + acStrToStr( file ) + " terminée." );
                
                // Réinitialisation du nombre d'objet
                obj = 0;
            }
            
            catch( ... )
            {
                print( "Impossible d'enregistrer les modifications." );
                acedSSFree( ssPoly );
                return;
            }
        }
        
        // Si le fichier est .xlsx(excel)
        else if( ext == "xlsx" )
        {
            // Initialisation du fichier excel
            workbook wb;
            worksheet ws = wb.active_sheet();
            
            // Ajout des en-têtes dans le feuille excel
            ws.cell( 1, 1 ).value( "Handle" );
            ws.cell( 2, 1 ).value( "Calque" );
            ws.cell( 3, 1 ).value( "Couleur" );
            ws.cell( 4, 1 ).value( "Type de ligne" );
            ws.cell( 5, 1 ).value( "Echelle de type de ligne" );
            ws.cell( 6, 1 ).value( "Epaisseur de ligne" );
            ws.cell( 7, 1 ).value( "Transparence" );
            ws.cell( 8, 1 ).value( "Epaisseur" );
            ws.cell( 9, 1 ).value( "Materiau" );
            ws.cell( 10, 1 ).value( "Longueur 2d" );
            ws.cell( 11, 1 ).value( "Longueur 3d" );
            ws.cell( 12, 1 ).value( "Surface" );
            
            ProgressBar prog_xls = ProgressBar( _T( "Progression:" ), size );
            
            // Iteration sur les objets
            int xls_it = 0;
            
            // Les entités polylines
            AcDbPolyline* pl2d = NULL;
            AcDb3dPolyline* pl3d = NULL;
            AcDbEntity* ent = NULL;
            
            AcDbObjectId obj_id;
            
            for( long i = 0; i < size; i++ )
            {
            
                // Le pointeur sur la polyligne
                obj_id = getObIdFromSs( ssPoly, i );
                
                //Ouvrir l'entité
                acdbOpenAcDbEntity( ent, obj_id, AcDb::kForRead );
                
                //verifier ent
                if( !ent )
                    continue;
                    
                if( ent->isKindOf( AcDbPolyline::desc() ) )
                {
                
                    //Recuperer la polyline
                    pl2d = AcDbPolyline::cast( ent );
                    
                    //veririfier pl2d
                    if( !pl2d )
                    {
                        prog_xls.moveUp( i );
                        ent->close();
                        continue;
                    }
                    
                    if( !isClosed( pl2d ) )
                    {
                    
                        prog_xls.moveUp( i );
                        pl2d->close();
                        continue;
                    }
                    
                    //Exporter l'entité
                    exportClosedPoly( ent, pData );
                }
                
                else if( ent->isKindOf( AcDb3dPolyline::desc() ) )
                {
                    //Recuperer la polyline
                    pl3d = AcDb3dPolyline::cast( ent );
                    
                    //veririfier pl2d
                    if( !pl3d )
                    {
                        prog_xls.moveUp( i );
                        ent->close();
                        continue;
                    }
                    
                    if( !isClosed( pl3d ) )
                    {
                        prog_xls.moveUp( i );
                        pl3d->close();
                        continue;
                    }
                    
                    //Exporter l'entité
                    exportClosedPoly( ent, pData );
                }
                
                else
                {
                    prog_xls.moveUp( i );
                    ent->close();
                    continue;
                }
                
                // Mettre le format de toute les cellules en Texte
                ws.columns( true ).number_format( number_format::text() );
                
                // Iteration sur le tableau des données
                for( long j = 0; j < pData.size(); j++ )
                {
                    //Convertir l'encodage des strings
                    auto data = latin1_to_utf8( acStrToStr( pData[j] ) );
                    
                    // Ecriture des données dans le fichier
                    ws.cell( j + 1, xls_it + 2 ).value( data );
                }
                
                // Incrémentation du nombre d'objet
                obj++;
                xls_it++;
                
                // Vider le tableau des données
                pData.clear();
                
                // Fermer la polyligne 2d
                pl2d->close();
                prog_xls.moveUp( i );
            }
            
            try
            {
                // Sauvegarder le fichier excel .xlsx
                wb.save( filename );
                
                // Affichage dans la console que les propriétes sont bien exportés
                print( "Exportation de " + to_string( obj ) + " polyligne(s) dans " + acStrToStr( file ) + " terminée." );
                
                // Réinitialisation de obj
                obj = 0;
            }
            
            catch( ... )
            {
                print( "Impossible d'enregistrer les modifications." );
                acedSSFree( ssPoly );
                return;
            }
            
        }
        
        //Si le fichier est .shp(shape)
        else if( ext == "shp" )
        {
        
        
            //Lire le fichier de parametre
            vector<AcString> field = createClosedPolyField();
            writePrjOrNot( file );
            
            //Creer un fichier shp
            SHPHandle myShapeFile = SHPCreate( file, SHPT_POLYGONZ );
            
            //Creer un dbfFile
            DBFHandle dbfHandle = DBFCreate( file );
            
            //Creer les champs
            int fieldNumber = createFields( dbfHandle, field );
            
            // Les entités polylines
            AcDbPolyline* pl2d = NULL;
            AcDb3dPolyline* pl3d = NULL;
            AcDbEntity* ent = NULL;
            
            AcDbObjectId obj_id;
            
            
            // Initialistion de la barre de progression
            ProgressBar prog = ProgressBar( _T( "Progression" ), size );
            
            //nombre de polylignes exportées
            int nbPoly = 0;
            
            //Boucler sur la selection
            for( int i = 0; i < size; i++ )
            {
                // Le pointeur sur la polyligne
                obj_id = getObIdFromSs( ssPoly, i );
                
                //Ouvrir l'entité
                acdbOpenAcDbEntity( ent, obj_id, AcDb::kForRead );
                
                //verifier ent
                if( !ent )
                    continue;
                    
                if( ent->isKindOf( AcDbPolyline::desc() ) )
                {
                
                    //Recuperer la polyline
                    pl2d = AcDbPolyline::cast( ent );
                    
                    //veririfier pl2d
                    if( !pl2d )
                    {
                        prog.moveUp( i );
                        ent->close();
                        continue;
                    }
                    
                    if( !isClosed( pl2d ) )
                    {
                    
                        prog.moveUp( i );
                        pl2d->close();
                        continue;
                    }
                    
                    //recuperer le nombre de sommets
                    int verTnumber = pl2d->numVerts();
                    
                    vector<AcGePoint3d> point;
                    AcGePoint3d p;
                    
                    //Boucler sur les vecteurs
                    for( int k = 0; k < verTnumber; k++ )
                    {
                        pl2d->getPointAt( k, p );
                        point.push_back( p );
                    }
                    
                    //Recuperer le point sur le sommet numero
                    pl2d->getPointAt( 0, p );
                    point.push_back( p );
                    
                    //Reciperer l'objet ligne
                    drawClosedPolyShp( pl2d, point, myShapeFile, dbfHandle, field );
                    
                    //incrementer le nombre de polylignes éxportées
                    nbPoly++;
                }
                
                else if( ent->isKindOf( AcDb3dPolyline::desc() ) )
                {
                    //Recuperer la polyline
                    pl3d = AcDb3dPolyline::cast( ent );
                    
                    //veririfier pl2d
                    if( !pl3d )
                    {
                        prog.moveUp( i );
                        ent->close();
                        continue;
                    }
                    
                    if( !isClosed( pl3d ) )
                    {
                        prog.moveUp( i );
                        pl3d->close();
                        continue;
                    }
                    
                    vector<AcGePoint3d> point;
                    
                    AcGePoint3d p;
                    //Boucer sur les sommets
                    AcDbObjectIterator* iter = pl3d->vertexIterator();
                    AcDb3dPolylineVertex* vertex;
                    
                    for( iter->start(); !iter->done(); iter->step() )
                    {
                        if( !pl3d->openVertex( vertex, iter->objectId(), AcDb::kForRead ) )
                        {
                            p = vertex->position();
                            point.push_back( p );
                        }
                    }
                    
                    vertex->close();
                    //Recuperer le point sur le sommet numero 0
                    pl3d->getStartPoint( p );
                    point.push_back( p );
                    
                    //Reciperer l'objet ligne
                    drawClosedPolyShp( pl3d, point, myShapeFile, dbfHandle, field );
                    
                    //incrementer le nombre de polylignes éxportées
                    nbPoly++;
                }
                
                else
                {
                    prog.moveUp( i );
                    ent->close();
                    continue;
                }
                
                
                Acad::ErrorStatus es = ent->close();
                
                // incrementer la barre de progression
                prog.moveUp( i );
            }
            
            DBFClose( dbfHandle );
            SHPClose( myShapeFile );
            
            // Affichage dans la console que les propriétes sont bien exportés
            print( "Exportation de " + to_string( nbPoly ) + " polyligne(s) dans " + acStrToStr( file ) + " terminée." );
            
        }
        
    }
    
    // Si la sélection est vide
    else
    {
        // Affichage dans la console
        print( "Aucune polyligne sélectionnée" );
    }
    
    
    acedSSFree( ssPoly );
}

void cmdExportLine()
{
    // Nom de la polyligne 2d
    ads_name ssLine;
    
    // Le fichier
    AcString file,
             // L'extension du fichier
             ext;
             
    file = _T( "" );
    
    
    //Recuperer le  chemin du dwg courant
    AcString current_folder = getCurrentFileFolder();
    current_folder += _T( "\\" );
    
    // Déclarer un AcDbLine
    AcDbLine* n_line = NULL;
    
    // Tableau de stockage des données des entités
    AcStringArray pData;
    
    // Demander à l'utilisateur de séléctionner les lines
    print( "Veuillez sélectionner les lines : " );
    
    // Les sélection sur les polylignes
    long size = getSsLine( ssLine, _T( "" ) );
    
    // Verification si la sélection n'est pas vide
    if( size != 0 )
    {
        // Nombre d'objet sélectionnées
        int obj = 0;
        
        // Demander à l'utilisateur le repertoire et l'extension du fichier à enregistrer
        file = askForFilePath( false,
                "xlsx;shp;csv;txt",
                "Enregistrer sous",
                current_folder );
                
        //Gestion d'erreur
        if( file == _T( "" ) )
        {
            print( "Commande annulée." );
            acedSSFree( ssLine );
            return;
        }
        
        // Prendre l'extension fu fichier
        ext = getFileExt( file );
        
        //Changer en utf8 l'encodage du chemin vers le fichier excel
        auto filename = latin1_to_utf8( acStrToStr( file ) );
        
        // Si le fichier est .txt ou .csv
        if( ext == "txt" || ext == "csv" )
        {
            // Flux de fichier
            fstream fichier( acStrToStr( file ), ios::app );
            
            // Ecriture des en-têtes du fichier
            fichier << "Handle;Calque;Couleur;Type de ligne;Echelle de type de ligne;Epaisseur de ligne;Transparence;Epaisseur;Materiau;Longueur 2d;Depart_X;Depart_Y;Depart_Z;Fin_X;Fin_Y;Fin_Z" << endl;
            
            // Iteration sur les objets
            for( long i = 0; i < size; i++ )
            {
                // Le pointeur sur la polyligne
                n_line = getLineFromSs( ssLine, i, AcDb::kForRead );
                
                // Appel de la fonction de traitement
                exportLines( n_line, pData );
                
                // Iteration sur le tableau des données
                for( long j = 0; j < pData.size(); j++ )
                {
                    // Ecriture des données dans le fichier
                    fichier << pData[j] << ";";
                }
                
                // Se mettre à la ligne
                fichier << endl;
                
                // Incrémentation du nomber d'objet
                obj++;
                
                // Vider le tableau des données
                pData.clear();
                
                // Fermer la polyligne
                n_line->close();
            }
            
            try
            {
            
                // Fermer le flux de fichier
                fichier.close();
                
                // Affichage dans la console que les informations sont bien exportés
                print( "Exportation de " + to_string( obj ) + "  lines dans " + acStrToStr( file ) + " terminée." );
                
                // Réinitialisation du nombre d'objet
                obj = 0;
                
            }
            
            catch( ... )
            {
                print( "Impossible d'enregistrer les modifications." );
                acedSSFree( ssLine );
                return;
            }
        }
        
        // Si le fichier est .xlsx(excel)
        else if( ext == "xlsx" )
        {
            // Initialisation du fichier excel
            workbook wb;
            worksheet ws = wb.active_sheet();
            
            // Ajout des en-têtes dans le feuille excel
            ws.cell( 1, 1 ).value( "Handle" );
            ws.cell( 2, 1 ).value( "Calque" );
            ws.cell( 3, 1 ).value( "Couleur" );
            ws.cell( 4, 1 ).value( "Type de ligne" );
            ws.cell( 5, 1 ).value( "Echelle de type de ligne" );
            ws.cell( 6, 1 ).value( "Epaisseur de ligne" );
            ws.cell( 7, 1 ).value( "Transparence" );
            ws.cell( 8, 1 ).value( "Epaisseur" );
            ws.cell( 9, 1 ).value( "Materiau" );
            ws.cell( 10, 1 ).value( "Longueur 2d" );
            ws.cell( 11, 1 ).value( "Depart_X" );
            ws.cell( 12, 1 ).value( "Depart_Y" );
            ws.cell( 13, 1 ).value( "Depart_Z" );
            ws.cell( 14, 1 ).value( "Fin_X" );
            ws.cell( 15, 1 ).value( "Fin_Y" );
            ws.cell( 16, 1 ).value( "Fin_Z" );
            
            
            // Iteration sur les objets
            for( long i = 0; i < size; i++ )
            {
                // Le pointeur sur la polyligne
                n_line = getLineFromSs( ssLine, i, AcDb::kForRead );
                
                // Safeguard
                if( !n_line )
                    continue;
                    
                // Appel de la fonction de traitement
                exportLines( n_line, pData );
                
                // Mettre le format de toute les cellules en Texte
                ws.columns( true ).number_format( number_format::text() );
                
                // Iteration sur le tableau des données
                for( long j = 0; j < pData.size(); j++ )
                {
                    // Ecriture des données dans le fichier
                    ws.cell( j + 1, i + 2 ).value( pData[j] );
                }
                
                // Incrémentation du nombre d'objet
                obj++;
                
                // Vider le tableau des données
                pData.clear();
                
                // Fermer la polyligne 2d
                n_line->close();
            }
            
            try
            {
                // Sauvegarder le fichier excel .xlsx
                wb.save( filename );
                
                // Affichage dans la console que les propriétes sont bien exportés
                print( "Exportation de " + to_string( obj ) + " line(s) dans " + acStrToStr( file ) + " terminée." );
                
                // Réinitialisation de obj
                obj = 0;
            }
            
            catch( ... )
            {
                print( "Impossible d'enregistrer les modifications." );
                acedSSFree( ssLine );
                return;
            }
        }
        
        //Si le fichier est .shp
        else if( ext == "shp" )
        {
            //Lire le fichier de parametre
            vector<AcString> field = createLineField();
            writePrjOrNot( file );
            
            //Creer un fichier shp
            SHPHandle myShapeFile = SHPCreate( file, SHPT_ARCZ );
            
            //Creer un dbfFile
            DBFHandle dbfHandle = DBFCreate( file );
            
            //Creer les champs
            int fieldNumber = createFields( dbfHandle, field );
            
            //Barre de progression
            ProgressBar prog = ProgressBar( _T( "Progression" ), size );
            
            //Nombre d'éléments convertis
            int nb_el = 0;
            
            //Boucler sur la selection
            for( int i = 0; i < size; i++ )
            {
                int vertNumber = 0;
                
                //Recuperer la i-eme line
                n_line = getLineFromSs( ssLine, i );
                
                //Verifier n_line
                if( !n_line )
                    continue;
                    
                //Recuperer l'objet ligne
                drawLineShp( n_line, myShapeFile, dbfHandle, field );
                
                //Liberer la mémoire
                n_line->close();
                
                //Progresser
                prog.moveUp( i );
                
                //Incrementer le nombre d'élements
                nb_el++;
            }
            
            //Liberer la mémoire
            DBFClose( dbfHandle );
            SHPClose( myShapeFile );
            
            // Message
            print( "Exportation de " + to_string( nb_el ) + " line(s) dans " + acStrToStr( file ) + " terminée." );
        }
        
    }
    
    // Si la sélection est vide
    else
    {
        // Affichage dans la console
        print( "Aucune line sélectionnée." );
    }
    
    // Liberer la sélection
    acedSSFree( ssLine );
    
}

void cmdExportCurve()
{

    //Définir la selection
    ads_name ssCurve;
    
    //Fichier & extension
    AcString file = _T( "" );
    AcString ext;
    
    //Recuperer le  chemin du dwg courant
    AcString current_folder = getCurrentFileFolder();
    current_folder += _T( "\\" );
    
    //Tableau de stockage de données
    AcStringArray pData;
    
    
    //Demander au dessinateur de faire une selection
    print( "Veuillez séléctionner la/les curve(s) : " );
    
    //Nombre de points dans la selection
    long size = getSelectionSet( ssCurve, _T( "" ), "LWPOLYLINE,POLYLINE,CIRCLE,LINE" );
    
    //Verifier le nombre de points
    if( size != 0 )
    {
    
        //Nombre d'objet sélectionné
        int obj = 0;
        
        //Initialiser file
        file = _T( "" );
        
        //Ou enregistrer le fichier?
        file = askForFilePath( false,
                "xlsx;shp;csv;txt",
                "Enregistrer sous",
                current_folder );
                
        //Gestion d'erreur
        if( file == _T( "" ) )
        {
            print( "Commande annulée." );
            acedSSFree( ssCurve );
            return;
        }
        
        //Recuperer l'extension du fichier
        ext = getFileExt( file );
        
        
        //Changer l'encodage du chemin vers le fichier excel en utf8
        auto filename = latin1_to_utf8( acStrToStr( file ) );
        
        //Si : .txt, .csv
        if( ext == "csv" || ext == "txt" )
        {
            //Ouvrir le fichier
            fstream fichier( acStrToStr( file ), ios::app );
            
            //Ecrire les en-têtes du fichier
            fichier << "Handle;Calque;Couleur;Type de ligne;Echelle de type de ligne;Epaisseur de ligne;Transparence;Epaisseur;Materiau;Longueur 2d;Longueur 3d;Fermée" << endl;
            
            //Déclarer un AcDbPoint
            AcDbObjectId obj_id;
            AcDbEntity* ent = NULL;
            
            //Barre de progression
            ProgressBar prog_txt = ProgressBar( _T( "Progression:" ), size );
            
            //Iterer sur la selection de points
            for( long i = 0; i < size; i++ )
            {
            
                //Le pointeur sur le point
                obj_id = getObIdFromSs( ssCurve, i );
                
                //Ouvrir l'entité
                acdbOpenAcDbEntity( ent, obj_id, AcDb::kForRead );
                
                //verifier ent
                if( !ent )
                    continue;
                    
                // Appel de la fonction de traitement
                exportCurve( ent, pData );
                
                
                // Iteration sur le tableau des données
                for( long j = 0; j < pData.size(); j++ )
                {
                    //Ecriture des données dans le fichier sélectionné
                    fichier << pData[j] << ";";
                }
                
                // Se mettre à la ligne
                fichier << endl;
                
                //Incrémentation du nombre d'objet
                obj++;
                
                //Reset pData
                pData.clear();
                
                //Liberer la mémoire
                ent->close();
                
                //Progresser
                prog_txt.moveUp( i );
            }
            
            try
            {
                //Fermeture du fichier
                fichier.close();
                
                //Affichage dans la console que les informations sont bien exportés
                print( "Exportation de " + to_string( obj ) + " curve(s) dans " + acStrToStr( file ) + " terminée." );
                
            }
            
            catch( ... )
            {
                print( "Impossible d'enregistrer le fichier." );
                acedSSFree( ssCurve );
                return;
            }
        }
        
        //Fichier excel
        else if( ext == "xlsx" )
        {
            //Initialisation du fichier excel
            workbook wb2;
            worksheet ws = wb2.active_sheet();
            
            //Ajout des en-tête du feuille
            ws.cell( 1, 1 ).value( u8"Handle" );
            ws.cell( 2, 1 ).value( u8"Calque" );
            ws.cell( 3, 1 ).value( u8"Couleur" );
            ws.cell( 4, 1 ).value( u8"Type de ligne" );
            ws.cell( 5, 1 ).value( u8"Echelle de type de ligne" );
            ws.cell( 6, 1 ).value( u8"Epaisseur de ligne" );
            ws.cell( 7, 1 ).value( u8"Transparence" );
            ws.cell( 8, 1 ).value( u8"Epaisseur" );
            ws.cell( 9, 1 ).value( u8"Materiau" );
            ws.cell( 10, 1 ).value( u8"Longueur 2d" );
            ws.cell( 11, 1 ).value( u8"Longueur 3d" );
            ws.cell( 12, 1 ).value( u8"Fermée" );
            
            //Déclarer un AcDbPoint
            AcDbObjectId obj_id;
            AcDbEntity* ent = NULL;
            
            //BArre de progression
            ProgressBar prog = ProgressBar( _T( "Progression" ), size );
            
            int cell_cp = 0;
            
            for( long i = 0; i < size; i++ )
            {
            
                //Recuperer le i-eme objectid
                obj_id = getObIdFromSs( ssCurve, i );
                
                //Ouvrir l'entité
                acdbOpenAcDbEntity( ent, obj_id, AcDb::kForRead );
                
                //verifier ent
                if( !ent )
                    continue;
                    
                // Appel de la fonction de traitement
                exportCurve( ent, pData );
                
                
                // Mettre le format des cellules en texte
                ws.columns( true ).number_format( number_format::text() );
                
                // Iteration sur les données
                for( long k = 0; k < pData.size(); k++ )
                {
                    auto data = latin1_to_utf8( acStrToStr( pData[k] ) );
                    //Ecriture dans le fichier excel
                    ws.cell( k + 1, cell_cp + 2 ).value( data );
                }
                
                //Incrémentation du nombre d'objet
                obj++;
                cell_cp++;
                
                // Vider le tableau des données
                pData.clear();
                
                //Fermer le point
                ent->close();
                
                //Progresser
                prog.moveUp( i );
            }
            
            try
            {
                //Enregistrer les modifications
                wb2.save( filename );
                
                //Affichage dans la console que les informations sont bien exportés
                print( "Exportation de " + to_string( obj ) + " curve(s) dans " + acStrToStr( file ) + " terminée." );
                
                //Réinitialisation de la valeur de obj
                obj = 0;
            }
            
            catch( ... )
            {
                print( "Impossible d'enregistrer le fichier." );
                acedSSFree( ssCurve );
                return;
            }
        }
        
        //Fichier shape
        else if( ext == "shp" )
        {
            //Lire le fichier de parametre
            vector<AcString> field = createCurveField();
            writePrjOrNot( file );
            
            //Creer un fichier shp
            SHPHandle myShapeFile = SHPCreate( file, SHPT_POINTZ );
            
            //Déclarer un AcDbPoint
            AcDbObjectId obj_id;
            AcDbEntity* ent = NULL;
            
            //Creer un dbfFile
            DBFHandle dbfHandle = DBFCreate( file );
            
            //Creer les champs
            int fieldNumber = createFields( dbfHandle, field );
            
            // Initialistion de la barre de progression
            ProgressBar prog = ProgressBar( _T( "Progression" ), size );
            
            //Nombre de points exportés
            int nb_el = 0;
            
            //Boucler sur toute les blocs
            for( int i = 0; i < size; i++ )
            {
                //Recuperer le i-eme objectid
                obj_id = getObIdFromSs( ssCurve, i );
                
                //Ouvrir l'entité
                acdbOpenAcDbEntity( ent, obj_id, AcDb::kForRead );
                
                //verifier ent
                if( !ent )
                    continue;
                    
                //Exporter les points
                drawCurve( ent, myShapeFile, dbfHandle, field );
                
                //Incrementer le nombre de points éxportés
                nb_el++;
                
                //Liberer la mémoire
                ent->close();
                
                //Progresser
                prog.moveUp( i );
            }
            
            
            //Liberer la mémoire
            DBFClose( dbfHandle );
            SHPClose( myShapeFile );
            
            // Messages
            print( "Exportation de " + to_string( nb_el ) + " curve(s) dans " + acStrToStr( file ) + " terminée." );
        }
    }
    
    //Si il n'y a pas de point sélectionné
    else
        print( "Aucun point sélectionné" );
        
        
    //Liberer la mémoire
    acedSSFree( ssCurve );
    
}

void cmdExportCircle()
{
    // Nom de la polyligne 2d
    ads_name ssCircle;
    
    // Le fichier
    AcString file,
             // L'extension du fichier
             ext;
             
    file = _T( "" );
    
    //Recuperer le  chemin du dwg courant
    AcString current_folder = getCurrentFileFolder();
    current_folder += _T( "\\" );
    
    // Tableau de stockage des données des entités
    AcStringArray pData;
    
    // Demander à l'utilisateur de séléctionner les lines
    print( "Veuillez faire une selection sur le dessin : " );
    
    // L'entité line
    AcDbCircle* p_circle = NULL;
    
    // Les sélection sur les polylignes
    long size = getSsCircle( ssCircle, _T( "" ) );
    
    // Verification si la sélection n'est pas vide
    if( size != 0 )
    {
        // Nombre d'objet sélectionnées
        int obj = 0;
        
        // Demander à l'utilisateur le repertoire et l'extension du fichier à enregistrer
        file = askForFilePath( false,
                "xlsx;shp;csv;txt",
                "Enregistrer sous",
                current_folder );
                
        //Gestion d'erreur
        if( file == _T( "" ) )
        {
            print( "Commande annulée." );
            acedSSFree( ssCircle );
            return;
        }
        
        // Prendre l'extension fu fichier
        ext = getFileExt( file );
        
        //Gestion d'erreur
        if( file == _T( "" ) )
        {
            print( "Commande annulée." );
            acedSSFree( ssCircle );
            return;
        }
        
        //Encoder en utf8 le chemin vers le fichier excel
        auto filename = latin1_to_utf8( acStrToStr( file ) );
        
        // Si le fichier est .txt ou .csv
        if( ext == "txt" || ext == "csv" )
        {
            // Flux de fichier
            fstream fichier( acStrToStr( file ), ios::app );
            
            // Ecriture des en-têtes du fichier
            fichier << "Handle;Calque;Couleur;Type de ligne;Echelle de type de ligne;Epaisseur de ligne;Transparence;Epaisseur;Materiau;X;Y;Z;Rayon" << endl;
            
            // Iteration sur les objets
            for( long i = 0; i < size; i++ )
            {
                // Le pointeur sur la polyligne
                p_circle = getCircleFromSs( ssCircle, i, AcDb::kForRead );
                
                // Appel de la fonction de traitement
                exportCircles( p_circle, pData );
                
                //Exclure la surface
                pData.removeLast();
                
                // Iteration sur le tableau des données
                for( long j = 0; j < pData.size(); j++ )
                {
                    // Ecriture des données dans le fichier
                    fichier << pData[j] << ";";
                }
                
                // Se mettre à la ligne
                fichier << endl;
                
                // Incrémentation du nomber d'objet
                obj++;
                
                // Vider le tableau des données
                pData.clear();
                
                // Fermer la polyligne
                p_circle->close();
            }
            
            try
            {
                // Fermer le flux de fichier
                fichier.close();
                
                // Affichage dans la console que les informations sont bien exportés
                print( "Exportation de " + to_string( obj ) + "  cercles dans " + acStrToStr( file ) + " terminée." );
                
            }
            
            catch( ... )
            {
                print( "Impossible d'enregistrer le fichier." );
                acedSSFree( ssCircle );
                return;
            }
        }
        
        // Si le fichier est .xlsx(excel)
        else if( ext == "xlsx" )
        {
            // Initialisation du fichier excel
            workbook wb;
            worksheet ws = wb.active_sheet();
            
            // Ajout des en-têtes dans le feuille excel
            ws.cell( 1, 1 ).value( "Handle" );
            ws.cell( 2, 1 ).value( "Calque" );
            ws.cell( 3, 1 ).value( "Couleur" );
            ws.cell( 4, 1 ).value( "Type de ligne" );
            ws.cell( 5, 1 ).value( "Echelle de type de ligne" );
            ws.cell( 6, 1 ).value( "Epaisseur de ligne" );
            ws.cell( 7, 1 ).value( "Transparence" );
            ws.cell( 8, 1 ).value( "Epaisseur" );
            ws.cell( 9, 1 ).value( "Materiau" );
            ws.cell( 10, 1 ).value( "X" );
            ws.cell( 11, 1 ).value( "Y" );
            ws.cell( 12, 1 ).value( "Z" );
            ws.cell( 13, 1 ).value( "Rayon" );
            
            ProgressBar prog = ProgressBar( _T( "Progression :" ), size );
            
            // Iteration sur les objets
            for( long i = 0; i < size; i++ )
            {
                // Le pointeur sur la polyligne
                p_circle = getCircleFromSs( ssCircle, i, AcDb::kForRead );
                
                // Safeguard
                if( !p_circle )
                {
                    prog.moveUp( i );
                    continue;
                }
                
                // Appel de la fonction de traitement
                exportCircles( p_circle, pData );
                
                //Exclure la surface
                pData.removeLast();
                
                // Mettre le format de toute les cellules en Texte
                ws.columns( true ).number_format( number_format::text() );
                
                // Iteration sur le tableau des données
                for( long j = 0; j < pData.size(); j++ )
                {
                    // Ecriture des données dans le fichier
                    ws.cell( j + 1, i + 2 ).value( pData[j] );
                }
                
                // Incrémentation du nombre d'objet
                obj++;
                
                // Vider le tableau des données
                pData.clear();
                
                // Fermer la polyligne 2d
                p_circle->close();
                
                prog.moveUp( i );
            }
            
            try
            {
                // Sauvegarder le fichier excel .xlsx
                wb.save( filename );
                
                // Affichage dans la console que les propriétes sont bien exportés
                print( "Exportation de " + to_string( obj ) + " cercle(s) dans " + acStrToStr( file ) + " terminée." );
                
            }
            
            catch( ... )
            {
                print( "Impossible d'enregistrer le fichier." );
                acedSSFree( ssCircle );
                return;
            }
        }
        
        //Si le fichier est .shp
        else if( ext == "shp" )
        {
            //Lire le fichier de parametre
            vector<AcString> field = createCircleField();
            writePrjOrNot( file );
            
            //Creer un fichier shp
            SHPHandle myShapeFile = SHPCreate( file, SHPT_ARCZ );
            
            //Creer un dbfFile
            DBFHandle dbfHandle = DBFCreate( file );
            
            //Creer les champs
            int fieldNumber = createFields( dbfHandle, field );
            
            
            // Initialistion de la barre de progression
            ProgressBar prog = ProgressBar( _T( "Progression" ), size );
            
            //Nombre d'objets exportés
            int nb_el = 0;
            
            //Boucler sur la selection
            for( int i = 0; i < size; i++ )
            {
                int vertNumber = 0;
                //Recuperer l'entite
                p_circle = getCircleFromSs( ssCircle, i );
                
                //Verifier
                if( !p_circle )
                {
                    // incrementer la barre de progression
                    prog.moveUp( i );
                    continue;
                }
                
                //Reciperer l'objet ligne
                drawCircleShp( p_circle, myShapeFile, dbfHandle, field );
                
                //Fermer la polyline
                p_circle->close();
                
                // incrementer la barre de progression
                prog.moveUp( i );
                
                //Incrementer le nombre d'objet éxportés
                nb_el++;
            }
            
            //Liberer la mémoire
            DBFClose( dbfHandle );
            SHPClose( myShapeFile );
            
            //Affichage dans la console que les propriétes sont bien exportés
            print( "Exportation de " + to_string( nb_el ) + " cercle(s) dans " + acStrToStr( file ) + " terminée." );
        }
        
    }
    
    else
        print( "Aucun cercle selectionné" );
        
    // Liberer la sélection
    acedSSFree( ssCircle );
}

void cmdExportBlockPoly()
{
    //Declarer les selections
    ads_name ssBlock;
    ads_name ssPoly;
    ads_name ssError, sserror;
    
    //Fichier & Extension
    AcString file, ext;
    
    file = _T( "" );
    
    //Recuperer le  chemin du dwg courant
    AcString current_folder = getCurrentFileFolder();
    current_folder += _T( "\\" );
    
    //Stockage : propriétés & données
    AcStringArray pData;
    AcStringArray propName;
    
    //Listes blocks ignorés
    vector<AcDbObjectId> vec_error;
    int ignored_blk_nb = 0;
    
    //Vecteur temporaire: liste des attributs & liste des noms de propriete
    vector<AcString> nomAttributs, nomPropriete;
    
    //Vecteur de type non fixe pour contenir les valeurs des proprietes
    vector<void*> propValue;
    
    //Set : (Noms des attributs sans doublons & Noms des propriétes sans doublons & Noms de définition) pour toutes les entités
    set<AcString> sAttr, sProp;
    
    //Nombre de blocs dynamiques
    int dynBlocCount = 0;
    
    //Messagee
    print( "Veuillez sélectionner le/les block(s) : " );
    
    //Recuperer le nombre de blockes
    long size = getSsBlock( ssBlock, "", _T( "" ) );
    
    
    //message
    print( "Veuillez sélectionner la/les polyline(s)" );
    
    
    //Selectionner tous les polylines fermées et recuperes leur nombre
    long size_p2d = getSelectionSet( ssPoly, _T( "" ), "LWPOLYLINE,POLYLINE" );
    
    //Declarer un vecteur de polyline pour stocker les polylines entourant un block
    vector<AcDbEntity*> vecPoly;
    
    //Déclaration d'un AcDbPolyline
    AcDbObjectId obj_id;
    AcDbEntity* ent = NULL;
    
    //Verifier
    if( size != 0 )
    {
    
        //Nombre d'objet exporté
        int obj = 0;
        
        //Demander chemin du fichier vers quoi on exporte
        file = askForFilePath( false,
                "xlsx;shp;csv;txt",
                "Enregistrer sous",
                current_folder );
                
        //Gestion d'erreur
        if( file == _T( "" ) )
        {
            print( "Commande annulée." );
            acedSSFree( ssBlock );
            acedSSFree( ssPoly );
            return;
        }
        
        //Recuperer l'extension
        ext = getFileExt( file );
        
        //Changer l'encodage du chemin vers le fichier excel en utf8
        auto filename = latin1_to_utf8( acStrToStr( file ) );
        
        //Déclarer une reference de blocke
        AcDbBlockReference* blk;
        
        //Fichier : .txt ou .csv
        if( ext == "txt" || ext == "csv" )
        {
            //Ouvrir le fichier
            fstream fichier( acStrToStr( file ), ios::app );
            
            //Ecriture des en-têtes du fichier
            fichier << "Handle;Calque;Nom;Couleur;Type de ligne;Echelle de type de ligne;Epaisseur de ligne;Transparence;Materiau;Pos_X;Pox_Y;Pos_Z;Rotation;Handle poly;Longueur 2d;Longueur 3d;Surface";
            
            //Itération sur les objets
            for( long i = 0; i < size; i++ )
            {
                //Pointeur sur le bloc
                blk = getBlockFromSs( ssBlock, i );
                
                //Verifier
                if( !blk )
                    continue;
                    
                //Recuperer la taille
                int taille_attr = getAttributesNamesListOfBlockRef( blk ).size();
                
                //Iteration sur la taille de la liste des noms des attributs
                for( long j = 0; j < taille_attr; j++ )
                {
                    //Mettre les attributs dans le set
                    sAttr.insert( getAttributesNamesListOfBlockRef( blk )[j] );
                }
                
                //Si block dynamique
                if( isDynamicBlock( blk ) )
                {
                    //Inrementation du nombre de bloc dynamique
                    dynBlocCount++;
                    
                    //Prendre les listes des noms de proprietes et aussi des valeurs de proprietes
                    getBlockPropWithValueList( blk, propName, propValue );
                    
                    //Iteration sur la liste des noms des proprietes
                    for( long pn = 0; pn < propName.size(); pn++ )
                    {
                        //Mettre les noms des proprietes dans le set (sans doublons)
                        sProp.insert( propName[pn] );
                    }
                }
                
                //Fermeture du bloc
                blk->close();
            }
            
            //Itération sur les données sans doublons dans le |setAttr| et les mettre dans le vecteur temporaire
            for( set<AcString>::iterator t = sAttr.begin(); t != sAttr.end(); t++ )
            {
                nomAttributs.push_back( *t );
                //Completer les en-tête avec la liste sans doublons des ATTRIBUTS
                fichier << ";" << acStrToStr( *t );
            }
            
            //Iteration sur les données sans doublons dans le |sProp| et les mettre dans le vecteur temporaire
            for( set<AcString>::iterator r = sProp.begin(); r != sProp.end(); r++ )
            {
                //Mettre les noms de proprietés sans doublons dans un vecteur
                nomPropriete.push_back( *r );
                //Completer les en-têtes avec la liste sans doublons des PROPRIETES
                fichier << ";" << acStrToStr( *r );
            }
            
            //Se mettre à la ligne pour commencer à inserer les données
            fichier << endl;
            
            //Barre de progression
            ProgressBar prog = ProgressBar( _T( "Progression :" ), size );
            
            //Reiterer sur les blocks
            for( long i = 0; i < size; i++ )
            {
            
                //Pointeur sur le block
                blk = getBlockFromSs( ssBlock, i );
                
                //Verifier blk
                if( !blk )
                    continue;
                    
                //Recuperer les infos poly
                for( int k = 0; k < size_p2d; k++ )
                {
                    //Recuperer le k-ième objectid
                    obj_id = getObIdFromSs( ssPoly, k );
                    
                    //Ouvrir l'entité
                    acdbOpenAcDbEntity( ent, obj_id, AcDb::kForRead );
                    
                    //Verifier ent
                    if( !ent )
                        continue;
                        
                    //Caster l'entité en polyline
                    if( ent->isKindOf( AcDbPolyline::desc() ) )
                    {
                        AcDbPolyline* poly_2d = AcDbPolyline::cast( ent );
                        
                        //verifier poly_2d
                        if( !poly_2d )
                        {
                            ent->close();
                            continue;
                        }
                        
                        //verifier
                        if( !isClosed( poly_2d ) )
                        {
                            ent->close();
                            continue;
                        }
                        
                        //Recuperer la position du block courant
                        AcGePoint3d pos_3d = blk->position();;
                        
                        //Transformer le point 3d en 2d pour pouvoir s'en servir dans la fonction isPointInPoly(poly2d, pt2d)
                        AcGePoint2d pos = AcGePoint2d::kOrigin;
                        pos.x = pos_3d.x;
                        pos.y = pos_3d.y;
                        
                        //Booleen pour exprimer le resultat
                        bool inPoly = isPointInPoly( poly_2d, pos );
                        
                        //Si le block n'est pas dans la polyline, continuer
                        if( !inPoly )
                        {
                            ent->close();
                            continue;
                        }
                        
                        //Sinon stocker ce polyline dans vecPoly2d
                        vecPoly.emplace_back( ent );
                    }
                    
                    else if( ent->isKindOf( AcDb3dPolyline::desc() ) )
                    {
                        AcDb3dPolyline* poly_3d = AcDb3dPolyline::cast( ent );
                        
                        //verifier poly_2d
                        if( !poly_3d )
                        {
                            ent->close();
                            continue;
                        }
                        
                        //verifier
                        if( !isClosed( poly_3d ) )
                        {
                            ent->close();
                            continue;
                        }
                        
                        //Recuperer la position du block courant
                        AcGePoint3d pos_3d = blk->position();;
                        
                        //Booleen pour exprimer le resultat
                        bool inPoly = isPointInPoly( poly_3d, pos_3d );
                        
                        //Si le block n'est pas dans la polyline, continuer
                        if( !inPoly )
                        {
                            ent->close();
                            continue;
                        }
                        
                        //Sinon stocker ce polyline dans vecPoly2d
                        vecPoly.emplace_back( ent );
                    }
                    
                    // Fermer la polyline
                    ent->close();
                }
                
                //Si le block n'est pas entouré de/des polyline(s) : on continue avec un autre block
                if( vecPoly.size() == 0 )
                {
                    //Progresser
                    prog.moveUp( i );
                    
                    //Enregistrer le nombre de block ignorés
                    vec_error.emplace_back( blk->id() );
                    
                    ignored_blk_nb++;
                    
                    //Fermeture du block
                    blk->close();
                    continue;
                }
                
                //Declarer un identifiant unique
                AcDbHandle poly_handle;
                
                //Longueur de la polyline
                double poly_length;
                
                //Longueur de la polyline
                double poly_length3d;
                
                //Surface de la polyline
                double poly_area;
                
                //Surface test
                double area = 1000000;
                
                //Indice du polyline qui a la plus petite surface dans vecPoly2D
                int pos_vp = 0;
                
                //Sinon : on cherche la polyline la plus proche et ayant la plus petite surface
                if( vecPoly.size() == 1 )
                {
                    //S'il n'y a qu'une polyline : recuperer son handle
                    poly_handle = vecPoly[0]->handle();
                    
                    if( vecPoly[0]->isKindOf( AcDbPolyline::desc() ) )
                    {
                        AcDbPolyline* temp = AcDbPolyline::cast( vecPoly[0] );
                        
                        //Recuperer la longueur de la polyline
                        poly_length = getLength( temp );
                        
                        poly_length3d = poly_length;
                        
                        //Recuperer la surface de la polyline
                        temp->getArea( poly_area );
                        
                    }
                    
                    else if( vecPoly[0]->isKindOf( AcDb3dPolyline::desc() ) )
                    {
                        AcDb3dPolyline* temp = AcDb3dPolyline::cast( vecPoly[0] );
                        
                        //Recuperer la longueur de la polyline
                        poly_length3d = getLength( temp );
                        
                        //Recuperer la longueur 2d de la polyline
                        poly_length = get2DLengthPoly( temp );
                        
                        //Recuperer la surface de la polyline
                        temp->getArea( poly_area );
                    }
                    
                    //Fermer l'entité
                    vecPoly[0]->close();
                    
                }
                
                else if( vecPoly.size() > 1 )
                {
                    //On parcours les polylines et on prend celle qui a la plus petite surface
                    for( int vp = 0; vp < vecPoly.size(); vp++ )
                    {
                        //Stocker la surface de la polyline courante
                        double temp_area = 0;
                        
                        if( vecPoly[vp]->isKindOf( AcDbPolyline::desc() ) )
                        {
                            AcDbPolyline* temp = AcDbPolyline::cast( vecPoly[0] );
                            
                            //verifier temp
                            if( !temp )
                            {
                                vecPoly[vp]->close();
                                continue;
                            }
                            
                            //Recuperer la surface de la polyline
                            temp->getArea( temp_area );
                            
                        }
                        
                        else if( vecPoly[vp]->isKindOf( AcDb3dPolyline::desc() ) )
                        {
                            AcDb3dPolyline* temp = AcDb3dPolyline::cast( vecPoly[0] );
                            
                            //verifier temp
                            if( !temp )
                            {
                                vecPoly[vp]->close();
                                continue;
                            }
                            
                            //Recuperer la surface de la polyline
                            temp->getArea( temp_area );
                        }
                        
                        //Tester si elle est petite que surface test
                        if( temp_area <= area )
                        {
                            //Recuperer sa surface
                            area = temp_area;
                            
                            //Recuperer son indice
                            pos_vp = vp;
                        }
                        
                        //Fermer la polyline
                        vecPoly[vp]->close();
                    }
                    
                    //Recuperer l'identificateur du polyline qui a la plus petite surface
                    poly_handle = vecPoly[pos_vp]->handle();
                    
                    
                    if( vecPoly[pos_vp]->isKindOf( AcDbPolyline::desc() ) )
                    {
                        AcDbPolyline* temp = AcDbPolyline::cast( vecPoly[pos_vp] );
                        
                        //Recuperer la longueur de la polyline
                        poly_length = getLength( temp );
                        
                        poly_length3d = poly_length;
                        
                        //Recuperer la surface de la polyline
                        temp->getArea( poly_area );
                        
                        //Fermer la polyline
                        vecPoly[pos_vp]->close();
                        
                    }
                    
                    else if( vecPoly[0]->isKindOf( AcDb3dPolyline::desc() ) )
                    {
                        AcDb3dPolyline* temp = AcDb3dPolyline::cast( vecPoly[pos_vp] );
                        
                        //Recuperer la longueur de la polyline
                        poly_length3d = getLength( temp );
                        
                        //Recuperer la longueur 2d de la polyline
                        poly_length = get2DLengthPoly( temp );
                        
                        //Recuperer la surface de la polyline
                        temp->getArea( poly_area );
                        
                        //Fermer la polyline
                        vecPoly[pos_vp]->close();
                    }
                }
                
                // Appel de la fonction de traitement
                exportBlock( blk, pData );
                
                // Iteration sur les données
                for( long j = 0; j < pData.size(); j++ )
                {
                    //Ecrire les données
                    fichier << pData[j] << ";";
                }
                
                //Inserer l'identifiant unique de la polyline
                fichier << poly_handle.ascii() << ";";
                
                //Inserer la longueur de la polyline
                fichier << to_string( poly_length ) << ";";
                
                //Inserer la longueur de la polyline
                fichier << to_string( poly_length3d ) << ";";
                
                //Inserer la surface de la polyline
                fichier << to_string( poly_area ) << ";";
                
                //Rcuperer la taille de nomAttributs
                int nb_att = nomAttributs.size();
                
                //Itération sur les attributs
                for( long att = 0; att < nb_att; att++ )
                {
                    // Si la valeur de l'attribut n'est pas vide
                    if( getAttributValue( blk, nomAttributs[att] ) != "empty" )
                    {
                        //Convertir en std::string
                        auto attrib_val = acStrToStr( getAttributValue( blk, nomAttributs[att] ) );
                        
                        //Ecrire la valeur de l'attribut
                        fichier << attrib_val << ";";
                    }
                    
                    // Si la valeur de l'attribut est vide
                    else
                    {
                        // Remplir la colonne par des N/A
                        fichier << "N/A;";
                    }
                }
                
                //Check si c'est un bloc dynamique
                if( isDynamicBlock( blk ) )
                {
                
                    //Iteration sur les noms de proprieté sans doublons
                    for( long propName = 0; propName < nomPropriete.size(); propName++ )
                    {
                        //Recuperer la valeur du propriete
                        AcString bProp = getPropertyStrFromBlock( blk, nomPropriete[propName] );
                        
                        //Si bProp n'est pas vide
                        if( bProp != _T( "" ) )
                        {
                            //Recuperer la valeur du propriete
                            auto prop_val = acStrToStr( getPropertyStrFromBlock( blk, nomPropriete[propName] ) );
                            
                            //Ecrire dans le fichier
                            fichier << prop_val << ";";
                        }
                        
                        // Si la valeur du propriete n'est pas vide
                        else
                        {
                            // Si la valeur du propriete est un double
                            if( getPropertyFromBlock( blk, nomPropriete[propName] ) != -1 || getPropertyFromBlock( blk, nomPropriete[propName] ) != 0 )
                            {
                                //Recuperer la propriete
                                auto prop_value = to_string( getPropertyFromBlock( blk, nomPropriete[propName] ) );
                                
                                //Ecrire les valeurs des proprietés dans le fichier
                                fichier << prop_value << ";";
                            }
                            
                            // Si ce n'est pas les deux types
                            else
                                fichier << "N/A;";
                        }
                        
                    }
                    
                }
                
                // Si ce n'est pas un bloc dynamique
                else
                {
                    for( long propName = 0; propName < nomPropriete.size(); propName++ )
                        fichier << "N/A;";
                }
                
                // Se mettre à la ligne après une écriture d'une ligne d'information des objets
                fichier << endl;
                
                //Incrémentation de obj
                obj++;
                
                //Fermeture du block
                blk->close();
                
                // Vider le tableau des données
                pData.clear();
                
                //Vider vecPoly
                vecPoly.clear();
                
                //Progresser
                prog.moveUp( i );
            }
            
            //Recuperer le nombre de blocks ignorés
            int taille_blk_ignored = vec_error.size();
            
            if( taille_blk_ignored > 0 )
            {
                for( auto& obj_id : vec_error )
                {
                    acdbGetAdsName( ssError, obj_id );
                    acedSSAdd( ssError, sserror, sserror );
                    acedSSFree( ssError );
                }
            }
            
            try
            {
                //Fermeture du fichier
                fichier.close();
                
                //Affichage dans la console que les informations sont bien exportés
                print( "Exportation de " + to_string( obj ) + " block(s) dont " + to_string( dynBlocCount ) + " bloc(s) dynamique(s) dans " + acStrToStr( file ) + " terminée." );
                print( "Nombre de blocks ignorés : " + to_string( taille_blk_ignored ) );
                //Réinitialisation de obj
                obj = 0;
            }
            
            catch( ... )
            {
                print( "Impossible d'enregistrer le fichier." );
                acedSSFree( ssBlock );
                acedSSFree( ssPoly );
                return;
            }
            
        }
        
        // Si c'est un fichier .xlsx
        else if( ext == "xlsx" )
        {
        
            //Initialisation du fichier excel
            workbook wb;
            worksheet ws = wb.active_sheet();
            
            //Ajout des en-tête
            ws.cell( 1, 1 ).value( "Handle" );
            ws.cell( 2, 1 ).value( "Calque" );
            ws.cell( 3, 1 ).value( "Nom" );
            ws.cell( 4, 1 ).value( "Couleur" );
            ws.cell( 5, 1 ).value( "Type de ligne" );
            ws.cell( 6, 1 ).value( "Echelle de type de ligne" );
            ws.cell( 7, 1 ).value( "Epaisseur de ligne" );
            ws.cell( 8, 1 ).value( "Transparence" );
            ws.cell( 9, 1 ).value( "Materiau" );
            ws.cell( 10, 1 ).value( "X" );
            ws.cell( 11, 1 ).value( "Y" );
            ws.cell( 12, 1 ).value( "Z" );
            ws.cell( 13, 1 ).value( "Rotation" );
            ws.cell( 14, 1 ).value( "Handle poly" );
            ws.cell( 15, 1 ).value( "Longueur 2d" );
            ws.cell( 16, 1 ).value( "Longueur 3d" );
            ws.cell( 17, 1 ).value( "Surface" );
            
            //Itération sur les objets
            for( long i = 0; i < size; i++ )
            {
            
                //Pointeur sur le bloc
                blk = getBlockFromSs( ssBlock, i );
                
                //Verifier le blk
                if( !blk )
                    continue;
                    
                //Recuperer la taille
                int taille_attr = getAttributesNamesListOfBlockRef( blk ).size();
                
                //Iteration sur la taille de la liste des noms des attributs
                for( long j = 0; j < taille_attr; j++ )
                {
                    //Mettre les attributs dans le set
                    sAttr.insert( getAttributesNamesListOfBlockRef( blk )[j] );
                }
                
                //Si block dynamique
                if( isDynamicBlock( blk ) )
                {
                    //Inrementation du nombre de bloc dynamique
                    dynBlocCount++;
                    
                    //Prendre les listes des noms de proprietes et aussi des valeurs de proprietes
                    getBlockPropWithValueList( blk, propName, propValue );
                    
                    //Iteration sur la liste des noms des proprietes
                    for( long pn = 0; pn < propName.size(); pn++ )
                    {
                        //Mettre les noms des proprietes dans le set (sans doublons)
                        sProp.insert( propName[pn] );
                    }
                }
                
                //Fermeture du bloc
                blk->close();
            }
            
            //Iteration sur les données sans doublons dans le |setAttr| et les mettre dans le vecteur
            for( set<AcString>::iterator t = sAttr.begin(); t != sAttr.end(); t++ )
            {
                // Mettre les données sans doublons dans un vecteur
                nomAttributs.push_back( *t );
            }
            
            //Recuperer la taille de nomAttributs
            int ta = nomAttributs.size();
            
            //Completer les colonnes d'en tête
            for( int a = 0; a < ta; a++ )
            {
                // Completer les données dans le fichier excel
                ws.cell( 17 + ( a + 1 ), 1 ).value( nomAttributs[a] );
            }
            
            
            //Iteration sur les données sans doublons dans le |sProp| et les mettre dans le vecteur temporaire
            for( set<AcString>::iterator r = sProp.begin(); r != sProp.end(); r++ )
            {
                //Mettre les noms des proprietes sans doublons dans le vecteur
                nomPropriete.push_back( *r );
            }
            
            //Recuperer la taille du nomPropriete
            int tp = nomPropriete.size();
            
            //Recuperer la taille du nomAttributs
            int tt = nomAttributs.size();
            
            //Iteration sur les noms de propriete
            for( long p = 0; p < tp; p++ )
            {
                //Completer les en-têtes dans le fichier excel
                ws.cell( ( 17 + tt + ( p + 1 ) ), 1 ).value( nomPropriete[p] );
            }
            
            //Barre de progression
            ProgressBar prog_xls = ProgressBar( _T( "Progression :" ), size );
            
            //Indice de ligne
            int ro = 2;
            
            //Iterer sur les blockes
            for( long i = 0; i < size; i++ )
            {
            
                //Recuperer le i-eme block
                blk = getBlockFromSs( ssBlock, i, GcDb::kForRead );
                
                //Recuperer les infos poly
                for( int k = 0; k < size_p2d; k++ )
                {
                    //Recuperer le k-ième objectid
                    obj_id = getObIdFromSs( ssPoly, k );
                    
                    //Ouvrir l'entité
                    acdbOpenAcDbEntity( ent, obj_id, AcDb::kForRead );
                    
                    //Verifier ent
                    if( !ent )
                        continue;
                        
                    //Caster l'entité en polyline
                    if( ent->isKindOf( AcDbPolyline::desc() ) )
                    {
                        AcDbPolyline* poly_2d = AcDbPolyline::cast( ent );
                        
                        //verifier poly_2d
                        if( !poly_2d )
                        {
                            ent->close();
                            continue;
                        }
                        
                        //verifier
                        if( !isClosed( poly_2d ) )
                        {
                            ent->close();
                            continue;
                        }
                        
                        //Recuperer la position du block courant
                        AcGePoint3d pos_3d = blk->position();;
                        
                        //Transformer le point 3d en 2d pour pouvoir s'en servir dans la fonction isPointInPoly(poly2d, pt2d)
                        AcGePoint2d pos = AcGePoint2d::kOrigin;
                        pos.x = pos_3d.x;
                        pos.y = pos_3d.y;
                        
                        //Booleen pour exprimer le resultat
                        bool inPoly = isPointInPoly( poly_2d, pos );
                        
                        //Si le block n'est pas dans la polyline, continuer
                        if( !inPoly )
                        {
                            ent->close();
                            continue;
                        }
                        
                        //Sinon stocker ce polyline dans vecPoly2d
                        vecPoly.emplace_back( ent );
                    }
                    
                    else if( ent->isKindOf( AcDb3dPolyline::desc() ) )
                    {
                        AcDb3dPolyline* poly_3d = AcDb3dPolyline::cast( ent );
                        
                        //verifier poly_2d
                        if( !poly_3d )
                        {
                            ent->close();
                            continue;
                        }
                        
                        //verifier
                        if( !isClosed( poly_3d ) )
                        {
                            ent->close();
                            continue;
                        }
                        
                        //Recuperer la position du block courant
                        AcGePoint3d pos_3d = blk->position();;
                        
                        //Booleen pour exprimer le resultat
                        bool inPoly = isPointInPoly( poly_3d, pos_3d );
                        
                        //Si le block n'est pas dans la polyline, continuer
                        if( !inPoly )
                        {
                            ent->close();
                            continue;
                        }
                        
                        //Sinon stocker ce polyline dans vecPoly2d
                        vecPoly.emplace_back( ent );
                    }
                    
                    // Fermer la polyline
                    ent->close();
                }
                
                //Si le block n'est pas entouré de/des polyline(s) : on continue avec un autre block
                if( vecPoly.size() == 0 )
                {
                    //Progresser
                    prog_xls.moveUp( i );
                    
                    //Enregistrer le nombre de block ignorés
                    vec_error.emplace_back( blk->id() );
                    
                    ignored_blk_nb++;
                    
                    //Fermeture du block
                    blk->close();
                    continue;
                }
                
                //Declarer un identifiant unique
                AcDbHandle poly_handle;
                
                //Longueur de la polyline
                double poly_length;
                
                //Longueur de la polyline
                double poly_length3d;
                
                //Surface de la polyline
                double poly_area;
                
                //Surface test
                double area = 1000000;
                
                //Indice du polyline qui a la plus petite surface dans vecPoly2D
                int pos_vp = 0;
                
                //Sinon : on cherche la polyline la plus proche et ayant la plus petite surface
                if( vecPoly.size() == 1 )
                {
                    //S'il n'y a qu'une polyline : recuperer son handle
                    poly_handle = vecPoly[0]->handle();
                    
                    if( vecPoly[0]->isKindOf( AcDbPolyline::desc() ) )
                    {
                        AcDbPolyline* temp = AcDbPolyline::cast( vecPoly[0] );
                        
                        //Recuperer la longueur de la polyline
                        poly_length = getLength( temp );
                        
                        poly_length3d = poly_length;
                        
                        //Recuperer la surface de la polyline
                        temp->getArea( poly_area );
                        
                    }
                    
                    else if( vecPoly[0]->isKindOf( AcDb3dPolyline::desc() ) )
                    {
                        AcDb3dPolyline* temp = AcDb3dPolyline::cast( vecPoly[0] );
                        
                        //Recuperer la longueur de la polyline
                        poly_length3d = getLength( temp );
                        
                        //Recuperer la longueur 2d de la polyline
                        poly_length = get2DLengthPoly( temp );
                        
                        //Recuperer la surface de la polyline
                        temp->getArea( poly_area );
                    }
                    
                    //Fermer l'entité
                    vecPoly[0]->close();
                    
                }
                
                else if( vecPoly.size() > 1 )
                {
                    //On parcours les polylines et on prend celle qui a la plus petite surface
                    for( int vp = 0; vp < vecPoly.size(); vp++ )
                    {
                        //Stocker la surface de la polyline courante
                        double temp_area = 0;
                        
                        if( vecPoly[vp]->isKindOf( AcDbPolyline::desc() ) )
                        {
                            AcDbPolyline* temp = AcDbPolyline::cast( vecPoly[0] );
                            
                            //verifier temp
                            if( !temp )
                            {
                                vecPoly[vp]->close();
                                continue;
                            }
                            
                            //Recuperer la surface de la polyline
                            temp->getArea( temp_area );
                            
                        }
                        
                        else if( vecPoly[vp]->isKindOf( AcDb3dPolyline::desc() ) )
                        {
                            AcDb3dPolyline* temp = AcDb3dPolyline::cast( vecPoly[0] );
                            
                            //verifier temp
                            if( !temp )
                            {
                                vecPoly[vp]->close();
                                continue;
                            }
                            
                            //Recuperer la surface de la polyline
                            temp->getArea( temp_area );
                        }
                        
                        //Tester si elle est petite que surface test
                        if( temp_area <= area )
                        {
                            //Recuperer sa surface
                            area = temp_area;
                            
                            //Recuperer son indice
                            pos_vp = vp;
                        }
                        
                        //Fermer la polyline
                        vecPoly[vp]->close();
                    }
                    
                    //Recuperer l'identificateur du polyline qui a la plus petite surface
                    poly_handle = vecPoly[pos_vp]->handle();
                    
                    
                    if( vecPoly[pos_vp]->isKindOf( AcDbPolyline::desc() ) )
                    {
                        AcDbPolyline* temp = AcDbPolyline::cast( vecPoly[pos_vp] );
                        
                        //Recuperer la longueur de la polyline
                        poly_length = getLength( temp );
                        
                        poly_length3d = poly_length;
                        
                        //Recuperer la surface de la polyline
                        temp->getArea( poly_area );
                        
                        //Fermer la polyline
                        vecPoly[pos_vp]->close();
                        
                    }
                    
                    else if( vecPoly[0]->isKindOf( AcDb3dPolyline::desc() ) )
                    {
                        AcDb3dPolyline* temp = AcDb3dPolyline::cast( vecPoly[pos_vp] );
                        
                        //Recuperer la longueur de la polyline
                        poly_length3d = getLength( temp );
                        
                        //Recuperer la longueur 2d de la polyline
                        poly_length = get2DLengthPoly( temp );
                        
                        //Recuperer la surface de la polyline
                        temp->getArea( poly_area );
                        
                        //Fermer la polyline
                        vecPoly[pos_vp]->close();
                    }
                }
                
                // Definir les formats de toutes les cellules dans les colonnes en texte
                ws.columns( true ).number_format( xlnt::number_format::text() );
                
                //Recuperer les données
                exportBlock( blk, pData );
                
                
                //Ecrire les données
                for( int d = 0; d < pData.size(); d++ )
                {
                    //Convertir les encodages de strings ent utf-8
                    auto data = latin1_to_utf8( acStrToStr( pData[d] ) );
                    
                    //Inserer les informations
                    ws.cell( d + 1, ro ).value( data );
                }
                
                //Inserer l'identifiant unique de la polyline
                ws.cell( 14, ro ).value( poly_handle.ascii() );
                
                //Inserer la longueur de la polyline
                ws.cell( 15, ro ).value( to_string( poly_length ) );
                
                //Inserer la longueur de la polyline
                ws.cell( 16, ro ).value( to_string( poly_length3d ) );
                
                //Inserer la surface de la polyline
                ws.cell( 17, ro ).value( to_string( poly_area ) );
                
                //Iteration sur les attributs
                for( long att = 0; att < nomAttributs.size(); att++ )
                {
                    if( getAttributValue( blk, nomAttributs[att] ) != "empty" )
                        ws.cell( 17 + ( att + 1 ), ro ).value( getAttributValue( blk, nomAttributs[att] ) );
                    else
                        ws.cell( 17 + ( att + 1 ), ro ).value( "N/A" );
                }
                
                // Check si c'est un bloc dynamique
                if( isDynamicBlock( blk ) )
                {
                    for( long propName = 0; propName < nomPropriete.size(); propName++ )
                    {
                        /*
                            OBTENIR LES VALEURS DES PROPRIETES A L'AIDE DES NOMS DE PROPRIETES
                        */
                        
                        AcString nProp = getPropertyStrFromBlock( blk, nomPropriete[propName] );
                        
                        // Si la valeur du propriete est un AcString
                        if( nProp != _T( "" ) )
                        {
                            //Recuperer la taille de nomAttributs
                            int tai_n = nomAttributs.size();
                            
                            //Recuperer le nom de propriete
                            auto pro_name = latin1_to_utf8( acStrToStr( getPropertyStrFromBlock( blk, nomPropriete[propName] ) ) );
                            
                            // Completer les colonnes excel avec les valeurs de proprietes
                            ws.cell( ( 17 + tai_n + 1 + propName ), ro ).value( pro_name );
                        }
                        
                        // Si la valeur du propriete n'est pas un AcString
                        else
                        {
                            // Si la valeur est un double
                            if( getPropertyFromBlock( blk, nomPropriete[propName] ) != -1 || getPropertyFromBlock( blk, nomPropriete[propName] ) != 0 )
                            {
                                // Ecrire les valeurs des proprietes dans les colonnes correspondantes
                                ws.cell( ( 17 + nomAttributs.size() + 1 + propName ), ro ).value( latin1_to_utf8( to_string( getPropertyFromBlock( blk, nomPropriete[propName] ) ) ) );
                            }
                            
                            // Si ce n'est pas dans les deux cas
                            else
                            {
                                // Mettre N/A à la place
                                ws.cell( ( 17 + nomAttributs.size() + 1 + propName ), ro ).value( "N/A" );
                            }
                        }
                    }
                }
                
                // Si ce n'est pas un blooc dynamique
                else
                {
                    // Iteration sur les noms de proprietes
                    for( long pNm = 0; pNm < nomPropriete.size(); pNm++ )
                        // Completer par N/A
                        ws.cell( ( 17 + nomAttributs.size() + 1 + pNm ), ro ).value( "N/A" );
                }
                
                //Incrementation de obj
                obj++;
                
                //Vider vecPoly
                vecPoly.clear();
                
                //Fermeture du block
                blk->close();
                
                //Incrementer la ligne
                ro++;
                
                // Vider le tableau des données
                pData.clear();
                
                //Progresser
                prog_xls.moveUp( i );
            }
            
            
            //Recuperer le nombre de blocks ignorés
            int taille_blk_ignored = vec_error.size();
            
            if( taille_blk_ignored > 0 )
            {
                for( auto& obj_id : vec_error )
                {
                    acdbGetAdsName( ssError, obj_id );
                    acedSSAdd( ssError, sserror, sserror );
                    acedSSFree( ssError );
                }
            }
            
            try
            {
                //Sauvegarder le fichier excel
                wb.save( filename );
                
                //Affichage dans la console le nombre de bloc(s) ou on a exporter ses informations + le nombre de bloc(s) dynamique dans la selection
                print( "Exportation de " + to_string( obj ) + " block(s) dont " + to_string( dynBlocCount ) + " dynamique(s) dans " + acStrToStr( file ) + " terminée." );
                print( "Nombre de blocks ignorés : " + to_string( taille_blk_ignored ) );
                //Réinitialisation de obj
                obj = 0;
            }
            
            catch( ... )
            {
                print( "Impossible d'enregistrer le fichier." );
                acedSSFree( ssBlock );
                acedSSFree( ssPoly );
                return;
            }
            
        }
        
        //Si c'est un fichier .shp
        else if( ext == "shp" )
        {
            //Lire le fichier de parametre
            vector<AcString> field = createBlockPoly2dField();
            writePrjOrNot( file );
            
            //Recuperer les points du surfaces
            vector<AcGePoint3d> point;
            map<int, vector<AcGePoint3d>> mapPoint;
            
            //Creer un fichier shp
            SHPHandle myShapeFile = SHPCreate( file, SHPT_POLYGONZ );
            
            //Creer un dbfFile
            DBFHandle dbfHandle = DBFCreate( file );
            
            //Creer les champs
            int fieldNumber = createFields( dbfHandle, field );
            
            // Initialistion de la barre de progression
            ProgressBar prog = ProgressBar( _T( "Progression" ), size );
            
            double* x;    double* y;    double* z;
            int exported = 0;
            
            //Boucler sur la selection
            for( int i = 0; i < size; i++ )
            {
                //Recuperer les blocks
                AcDbBlockReference* block = getBlockFromSs( ssBlock, i, AcDb::kForRead );
                
                //Verifier block
                if( !block )
                    continue;
                    
                //surface index
                int iS; bool toDraw = false;
                int erea = INT32_MAX;
                
                //Recuperer son point d'insertion
                AcGePoint3d ptBlock = block->position();
                
                //Recuperer l'object id de la polyline
                AcDbObjectId idPoly;
                
                //Boucler sur le polyligne
                for( int j = 0; j < size_p2d; j++ )
                {
                    obj_id = getObIdFromSs( ssPoly, j );
                    
                    //Ouvrir l'entitié
                    acdbOpenAcDbEntity( ent, obj_id, AcDb::kForRead );
                    
                    //verifier ent
                    if( !ent )
                        continue;
                        
                    if( ent->isKindOf( AcDbPolyline::desc() ) )
                    {
                        //Recuperer le polyligne
                        AcDbPolyline* poly = AcDbPolyline::cast( ent );
                        
                        //Verifier poly
                        if( !poly )
                            continue;
                            
                        if( !isClosed( poly ) )
                        {
                            ent->close();
                            continue;
                        }
                        
                        //Recuperer le surface
                        double surface;
                        poly->getArea( surface );
                        surface = abs( surface );
                        
                        //Verifier si le block est a l'interieur du polyligne
                        if( isPointInPoly( poly, AcGePoint2d( ptBlock.x, ptBlock.y ) ) )
                        {
                            if( surface < erea )
                            {
                                erea = surface;
                                toDraw = true;
                                iS = j;
                                
                                if( mapPoint.find( j ) == mapPoint.end() )
                                {
                                    //Recuperer les vertexs
                                    int vertNumber = poly->numVerts();
                                    
                                    point.resize( 0 );
                                    AcGePoint3d p;
                                    
                                    //Boucler sur les vecteurs
                                    for( int k = 0; k < vertNumber; k++ )
                                    {
                                        poly->getPointAt( k, p );
                                        point.push_back( p );
                                    }
                                    
                                    //Recuperer le point sur le sommet numero
                                    poly->getPointAt( 0, p );
                                    point.push_back( p );
                                    
                                    //Recuperer l'object id de la polyline
                                    idPoly = poly->id();
                                    
                                    //Ajoute les infos dans la map
                                    mapPoint.insert( { j, point } );
                                }
                            }
                        }
                        
                        poly->close();
                    }
                    
                    else if( ent->isKindOf( AcDb3dPolyline::desc() ) )
                    {
                        //Recuperer le polyligne
                        AcDb3dPolyline* poly = AcDb3dPolyline::cast( ent );
                        
                        //Verifier poly
                        if( !poly )
                            continue;
                            
                        if( !isClosed( poly ) )
                        {
                            ent->close();
                            continue;
                        }
                        
                        //Recuperer le surface
                        double surface;
                        poly->getArea( surface );
                        surface = abs( surface );
                        
                        //Verifier si le block est a l'interieur du polyligne
                        if( isPointInPoly( poly, ptBlock ) )
                        {
                            if( surface < erea )
                            {
                                erea = surface;
                                toDraw = true;
                                iS = j;
                                
                                if( mapPoint.find( j ) == mapPoint.end() )
                                {
                                    //Reset point
                                    point.resize( 0 );
                                    
                                    //Nombre de sommets
                                    int vertNumber = getNumberOfVertex( poly );
                                    
                                    //Créer un itérateur de sommet
                                    AcDbObjectIterator* iterPoly = poly->vertexIterator();
                                    AcDb3dPolylineVertex* vertex;
                                    
                                    //Nombre totale des sommets
                                    vertNumber++;
                                    
                                    //Recuperer les sommets de la polyline
                                    for( iterPoly->start(); !iterPoly->done(); iterPoly->step() )
                                    {
                                        if( Acad::eOk == poly->openVertex( vertex, iterPoly->objectId(), AcDb::kForRead ) )
                                            point.emplace_back( vertex->position() );
                                            
                                    }
                                    
                                    //Liberer la mémoire
                                    vertex->close();
                                    
                                    //Recuperer le premier point
                                    AcGePoint3d p;
                                    poly->getStartPoint( p );
                                    
                                    //Rajouter le premier point
                                    point.emplace_back( p );
                                    
                                    //Recuperer l'object id de la polyline
                                    idPoly = poly->id();
                                    
                                    //Ajoute les infos dans la map
                                    mapPoint.insert( { j, point } );
                                }
                            }
                        }
                        
                        poly->close();
                    }
                }
                
                AcDbEntity* entPoly = NULL;
                
                if( toDraw )
                {
                    acdbOpenAcDbEntity( entPoly, idPoly, AcDb::kForRead );
                    
                    if( !entPoly )
                        continue;
                        
                    //Recuperer le vecteur point du rang du sommets
                    vector<AcGePoint3d> vecPoint = mapPoint.at( iS );
                    drawBlockPoly( block, entPoly, vecPoint, myShapeFile, dbfHandle, field );
                    exported++;
                    
                    entPoly->close();
                }
                
                else
                {
                
                    //Ajouter le block dans la selection courante
                    acdbGetAdsName( ssError, block->id() );
                    acedSSAdd( ssError, sserror, sserror );
                    acedSSFree( ssError );
                    
                    
                    //Incrementer le nombre de blocks ignorés
                    ignored_blk_nb++;
                }
                
                prog.moveUp( i );
                //Fermer le block
                block->close();
                
                // incrementer la barre de progression
                prog.moveUp( i );
                
                
                
            }
            
            DBFClose( dbfHandle );
            SHPClose( myShapeFile );
            
            //Message
            print( "Exportation de " + to_string( exported ) + " block(s) dans " + acStrToStr( file ) + " terminée." );
            print( "Nombre de blocks ignorés : " + to_string( ignored_blk_nb ) );
            
            
        }
        
    }
    
    else
        print( "Aucun block selectionné." );
        
        
    //Liberer la mémoire
    acedSSFree( ssBlock );
    acedSSFree( ssPoly );
    acedSSSetFirst( sserror, sserror );
}

void cmdExportBlockPoly2D()
{
    //  Nombre de blocs ignorés
    int ignoredBlock = 0;
    
    //Définition de la selection
    ads_name ssBlock;
    ads_name ssPoly2d;
    
    ads_name ssSelection, ssEnd;
    
    //  Vecteur de stockage des blocks
    vector<AcDbBlockReference*> v_blocks;
    
    //  Vecteur de stockage des polylignes 2d
    vector<AcDbPolyline*> v_poly2d;
    
    //Fichier & Extension
    AcString file, ext;
    
    file = _T( "" );
    
    //Recuperer le  chemin du dwg courant
    AcString current_folder = getCurrentFileFolder();
    current_folder += _T( "\\" );
    
    //Stockage : propriétés & données
    AcStringArray pData;
    AcStringArray propName;
    
    //Vecteur temporaire: liste des attributs & liste des noms de propriete
    vector<AcString> nomAttributs, nomPropriete;
    
    //Vecteur de type non fixe pour contenir les valeurs des proprietes
    vector<void*> propValue;
    
    //Set : (Noms des attributs sans doublons & Noms des propriétes sans doublons & Noms de définition) pour toutes les entités
    set<AcString> sAttr, sProp;
    
    //Nombre de blocs dynamiques
    int dynBlocCount = 0;
    
    //  Demander la sélection de(s) block(s)
    print( "Veuillez sélectionner le(s) block(s) : " );
    
    //  Récupérer le nombre de block
    long size_blocks = getSsBlock( ssBlock, "" );
    
    //  Demander la sélection de(s) polyligne(s)
    print( "Veuillez sélectionner le(s) polyligne(s) : " );
    
    //  Récupérer le nombre de polyligne
    long size_poly = getSsPoly2D( ssPoly2d, "" );
    
    //Declarer un vecteur de polyline pour stocker les polylines entourant un block
    vector<AcDbPolyline*> vecPoly2d;
    
    //Déclaration d'un AcDbPolyline
    AcDbPolyline* poly_2d = NULL;
    
    //Verifier
    if( size_blocks != 0 && size_poly != 0 )
    {
        //Nombre d'objet exporté
        int obj = 0;
        
        //Demander chemin du fichier vers quoi on exporte
        file = askForFilePath( false,
                "xlsx;shp;csv;txt",
                "Enregistrer sous",
                current_folder );
                
        //Gestion d'erreur
        if( file == _T( "" ) )
        {
            print( "Commande annulée." );
            acedSSFree( ssBlock );
            acedSSFree( ssPoly2d );
            return;
        }
        
        //Recuperer l'extension
        ext = getFileExt( file );
        
        //Changer l'encodage du chemin vers le fichier excel en utf8
        auto filename = latin1_to_utf8( acStrToStr( file ) );
        
        //Déclarer une reference de blocke
        AcDbBlockReference* blk;
        
        //  Itération sur les blocks
        for( size_t b = 0; b < size_blocks; b++ )
        {
            //  Prendre le i-ème block
            AcDbBlockReference* _blk = getBlockFromSs( ssBlock, b );
            
            //  Ajouter dans le vecteur de blocks
            v_blocks.emplace_back( _blk );
        }
        
        //  Itération sur les polylignes
        for( size_t p = 0; p < size_poly; p++ )
        {
            //  Prendre le i-ème polyligne
            AcDbPolyline* _poly = getPoly2DFromSs( ssPoly2d, p );
            
            //  Ajouter dans le vecteur de polyligne
            v_poly2d.emplace_back( _poly );
        }
        
        //Fichier : .txt ou .csv
        if( ext == "txt" || ext == "csv" )
        {
            //Ouvrir le fichier
            fstream fichier( acStrToStr( file ), ios::app );
            
            //Ecriture des en-têtes du fichier
            fichier << "Handle;Calque;Nom;Couleur;Type de ligne;Echelle de type de ligne;Epaisseur de ligne;Transparence;Materiau;X;Y;Z;Rotation;Handle poly;Longueur 2D;Surface";
            
            //Itération sur les blocks
            for( long i = 0; i < size_blocks; i++ )
            {
                //  Prendre le block
                blk = v_blocks[i];
                
                //Verifier
                if( !blk )
                    continue;
                    
                //Recuperer la taille
                int taille_attr = getAttributesNamesListOfBlockRef( blk ).size();
                
                //Iteration sur la taille de la liste des noms des attributs
                for( long j = 0; j < taille_attr; j++ )
                {
                    //Mettre les attributs dans le set
                    sAttr.insert( getAttributesNamesListOfBlockRef( blk )[j] );
                }
                
                //Si block dynamique
                if( isDynamicBlock( blk ) )
                {
                    //Inrementation du nombre de bloc dynamique
                    dynBlocCount++;
                    
                    //Prendre les listes des noms de proprietes et aussi des valeurs de proprietes
                    getBlockPropWithValueList( blk, propName, propValue );
                    
                    //Iteration sur la liste des noms des proprietes
                    for( long pn = 0; pn < propName.size(); pn++ )
                    {
                        //Mettre les noms des proprietes dans le set (sans doublons)
                        sProp.insert( propName[pn] );
                    }
                }
                
                //Fermeture du bloc
                blk->close();
            }
            
            //Itération sur les données sans doublons dans le |setAttr| et les mettre dans le vecteur temporaire
            for( set<AcString>::iterator t = sAttr.begin(); t != sAttr.end(); t++ )
            {
                nomAttributs.push_back( *t );
                //Completer les en-tête avec la liste sans doublons des ATTRIBUTS
                fichier << ";" << acStrToStr( *t );
            }
            
            //Iteration sur les données sans doublons dans le |sProp| et les mettre dans le vecteur temporaire
            for( set<AcString>::iterator r = sProp.begin(); r != sProp.end(); r++ )
            {
                //Mettre les noms de proprietés sans doublons dans un vecteur
                nomPropriete.push_back( *r );
                //Completer les en-têtes avec la liste sans doublons des PROPRIETES
                fichier << ";" << acStrToStr( *r );
            }
            
            //Se mettre à la ligne pour commencer à inserer les données
            fichier << endl;
            
            //Barre de progression
            ProgressBar prog = ProgressBar( _T( "Progression :" ), size_blocks );
            
            //Reiterer sur les blocks
            for( long i = 0; i < v_blocks.size(); i++ )
            {
                //Pointeur sur le block
                blk = v_blocks[i];
                
                //Verifier blk
                if( !blk )
                    continue;
                    
                //Recuperer les infos poly
                for( int k = 0; k < v_poly2d.size(); k++ )
                {
                    //Recuperer le k-ième polyline 2d
                    poly_2d = v_poly2d[k];
                    
                    //Verifier
                    if( !poly_2d )
                        continue;
                        
                    //Recuperer la position du block courant
                    AcGePoint3d pos_3d = blk->position();;
                    
                    //Transformer le point 3d en 2d pour pouvoir s'en servir dans la fonction isPointInPoly(poly2d, pt2d)
                    AcGePoint2d pos = AcGePoint2d::kOrigin;
                    pos.x = pos_3d.x;
                    pos.y = pos_3d.y;
                    
                    //Booleen pour exprimer le resultat
                    bool inPoly = isPointInPoly( poly_2d, pos );
                    
                    //Si le block n'est pas dans la polyline, continuer
                    if( !inPoly )
                    {
                        //  Fermer la polyligne et fermer
                        poly_2d->close();
                        continue;
                    }
                    
                    //Sinon stocker ce polyline dans vecPoly2d
                    vecPoly2d.emplace_back( poly_2d );
                    
                    // Fermer la polyline
                    poly_2d->close();
                }
                
                //Si le block n'est pas entouré de/des polyline(s) : on continue avec un autre block
                if( vecPoly2d.size() == 0 )
                {
                    //  Ajouter l'objet dans la sélection
                    acdbGetAdsName( ssSelection, blk->objectId() );
                    acedSSAdd( ssSelection, ssEnd, ssEnd );
                    acedSSFree( ssSelection );
                    
                    //  Incrémentation du nombre de block ignoré
                    ignoredBlock += 1;
                    
                    //Fermeture du block
                    blk->close();
                    
                    //  Continuer
                    continue;
                }
                
                //Declarer un identifiant unique
                AcDbHandle poly_handle;
                
                //Longueur de la polyline
                double poly_length;
                
                //Surface de la polyline
                double poly_area;
                
                //Surface test
                double area = 1000000;
                
                //Indice du polyline qui a la plus petite surface dans vecPoly2D
                int pos_vp = 0;
                
                //Sinon : on cherche la polyline la plus proche et ayant la plus petite surface
                if( vecPoly2d.size() == 1 )
                {
                    //S'il n'y a qu'une polyline : recuperer son handle
                    poly_handle = vecPoly2d[0]->handle();
                    
                    //Recuperer la longueur de la polyline
                    poly_length = getLength( vecPoly2d[0] );
                    
                    //Recuperer la surface de la polyline
                    vecPoly2d[0]->getArea( poly_area );
                }
                
                else if( vecPoly2d.size() > 1 )
                {
                    //On parcours les polylines et on prend celle qui a la plus petite surface
                    for( int vp = 0; vp < vecPoly2d.size(); vp++ )
                    {
                        //Stocker la surface de la polyline courante
                        double temp_area = 0;
                        vecPoly2d[vp]->getArea( temp_area );
                        
                        //Tester si elle est petite que surface test
                        if( temp_area <= area )
                        {
                            //Recuperer sa surface
                            area = temp_area;
                            
                            //Recuperer son indice
                            pos_vp = vp;
                        }
                        
                        //Fermer la polyline
                        vecPoly2d[vp]->close();
                    }
                    
                    //Recuperer l'identificateur du polyline qui a la plus petite surface
                    poly_handle = vecPoly2d[pos_vp]->handle();
                    
                    //Recuperer la longueur de la polyline
                    poly_length = getLength( vecPoly2d[pos_vp] );
                    
                    //Recuperer la surface de la polyline
                    vecPoly2d[pos_vp]->getArea( poly_area );
                }
                
                // Appel de la fonction de traitement
                exportBlock( blk, pData );
                
                // Iteration sur les données
                for( long j = 0; j < pData.size(); j++ )
                {
                    //Ecrire les données
                    fichier << pData[j] << ";";
                }
                
                //Inserer l'identifiant unique de la polyline
                fichier << poly_handle.ascii() << ";";
                
                //Inserer la longueur de la polyline
                fichier << to_string( poly_length ) << ";";
                
                //Inserer la surface de la polyline
                fichier << to_string( poly_area ) << ";";
                
                //Rcuperer la taille de nomAttributs
                int nb_att = nomAttributs.size();
                
                //Itération sur les attributs
                for( long att = 0; att < nb_att; att++ )
                {
                    // Si la valeur de l'attribut n'est pas vide
                    if( getAttributValue( blk, nomAttributs[att] ) != "empty" )
                    {
                        //Convertir en std::string
                        auto attrib_val = acStrToStr( getAttributValue( blk, nomAttributs[att] ) );
                        
                        //Ecrire la valeur de l'attribut
                        fichier << attrib_val << ";";
                    }
                    
                    // Si la valeur de l'attribut est vide
                    else
                    {
                        // Remplir la colonne par des N/A
                        fichier << "N/A;";
                    }
                }
                
                //Check si c'est un bloc dynamique
                if( isDynamicBlock( blk ) )
                {
                
                    //Iteration sur les noms de proprieté sans doublons
                    for( long propName = 0; propName < nomPropriete.size(); propName++ )
                    {
                        //Recuperer la valeur du propriete
                        AcString bProp = getPropertyStrFromBlock( blk, nomPropriete[propName] );
                        
                        //Si bProp n'est pas vide
                        if( bProp != _T( "" ) )
                        {
                            //Recuperer la valeur du propriete
                            auto prop_val = acStrToStr( getPropertyStrFromBlock( blk, nomPropriete[propName] ) );
                            
                            //Ecrire dans le fichier
                            fichier << prop_val << ";";
                        }
                        
                        // Si la valeur du propriete n'est pas vide
                        else
                        {
                            // Si la valeur du propriete est un double
                            if( getPropertyFromBlock( blk, nomPropriete[propName] ) != -1 || getPropertyFromBlock( blk, nomPropriete[propName] ) != 0 )
                            {
                                //Recuperer la propriete
                                auto prop_value = to_string( getPropertyFromBlock( blk, nomPropriete[propName] ) );
                                
                                //Ecrire les valeurs des proprietés dans le fichier
                                fichier << prop_value << ";";
                            }
                            
                            // Si ce n'est pas les deux types
                            else
                                fichier << "N/A;";
                        }
                    }
                }
                
                // Si ce n'est pas un bloc dynamique
                else
                {
                    for( long propName = 0; propName < nomPropriete.size(); propName++ )
                        fichier << "N/A;";
                }
                
                // Se mettre à la ligne après une écriture d'une ligne d'information des objets
                fichier << endl;
                
                //Incrémentation de obj
                obj++;
                
                //Fermeture du block
                blk->close();
                
                // Vider le tableau des données
                pData.clear();
                
                //Vider vecPoly
                vecPoly2d.clear();
                
                //Progresser
                prog.moveUp( i );
            }
            
            try
            {
                //Fermeture du fichier
                fichier.close();
                
                //Affichage dans la console que les informations sont bien exportés
                print( "Exportation de " + to_string( obj ) + " block(s) dont " + to_string( dynBlocCount ) + " bloc(s) dynamique(s) dans " + acStrToStr( file ) + " terminée." );
                
                //  Affichage dans la console du nombre de bloc ignoré
                print( "\nNombre de bloc(s) ignoré = " + to_string( ignoredBlock ) );
                
                //Réinitialisation de obj
                obj = 0;
            }
            
            catch( ... )
            {
                print( "Impossible d'enregistrer le fichier." );
                acedSSFree( ssBlock );
                acedSSFree( ssPoly2d );
                return;
            }
            
        }
        
        // Si c'est un fichier .xlsx
        else if( ext == "xlsx" )
        {
            //Initialisation du fichier excel
            workbook wb;
            worksheet ws = wb.active_sheet();
            
            //Ajout des en-tête
            ws.cell( 1, 1 ).value( "Handle" );
            ws.cell( 2, 1 ).value( "Calque" );
            ws.cell( 3, 1 ).value( "Nom" );
            ws.cell( 4, 1 ).value( "Couleur" );
            ws.cell( 5, 1 ).value( "Type de ligne" );
            ws.cell( 6, 1 ).value( "Echelle de type de ligne" );
            ws.cell( 7, 1 ).value( "Epaisseur de ligne" );
            ws.cell( 8, 1 ).value( "Transparence" );
            ws.cell( 9, 1 ).value( "Materiau" );
            ws.cell( 10, 1 ).value( "X" );
            ws.cell( 11, 1 ).value( "Y" );
            ws.cell( 12, 1 ).value( "Z" );
            ws.cell( 13, 1 ).value( "Rotation" );
            ws.cell( 14, 1 ).value( "Handle poly" );
            ws.cell( 15, 1 ).value( "Longueur 2D" );
            ws.cell( 16, 1 ).value( "Surface" );
            
            //Itération sur les objets
            for( long i = 0; i < size_blocks; i++ )
            {
                blk = v_blocks[i];
                
                //Verifier le blk
                if( !blk )
                    continue;
                    
                //Recuperer la taille
                int taille_attr = getAttributesNamesListOfBlockRef( blk ).size();
                
                //Iteration sur la taille de la liste des noms des attributs
                for( long j = 0; j < taille_attr; j++ )
                {
                    //Mettre les attributs dans le set
                    sAttr.insert( getAttributesNamesListOfBlockRef( blk )[j] );
                }
                
                //Si block dynamique
                if( isDynamicBlock( blk ) )
                {
                    //Inrementation du nombre de bloc dynamique
                    dynBlocCount++;
                    
                    //Prendre les listes des noms de proprietes et aussi des valeurs de proprietes
                    getBlockPropWithValueList( blk, propName, propValue );
                    
                    //Iteration sur la liste des noms des proprietes
                    for( long pn = 0; pn < propName.size(); pn++ )
                    {
                        //Mettre les noms des proprietes dans le set (sans doublons)
                        sProp.insert( propName[pn] );
                    }
                }
                
                //Fermeture du bloc
                blk->close();
            }
            
            //Iteration sur les données sans doublons dans le |setAttr| et les mettre dans le vecteur
            for( set<AcString>::iterator t = sAttr.begin(); t != sAttr.end(); t++ )
            {
                // Mettre les données sans doublons dans un vecteur
                nomAttributs.push_back( *t );
            }
            
            //Recuperer la taille de nomAttributs
            int ta = nomAttributs.size();
            
            //Completer les colonnes d'en tête
            for( int a = 0; a < ta; a++ )
            {
                // Completer les données dans le fichier excel
                ws.cell( 16 + ( a + 1 ), 1 ).value( nomAttributs[a] );
            }
            
            
            //Iteration sur les données sans doublons dans le |sProp| et les mettre dans le vecteur temporaire
            for( set<AcString>::iterator r = sProp.begin(); r != sProp.end(); r++ )
            {
                //Mettre les noms des proprietes sans doublons dans le vecteur
                nomPropriete.push_back( *r );
            }
            
            //Recuperer la taille du nomPropriete
            int tp = nomPropriete.size();
            
            //Recuperer la taille du nomAttributs
            int tt = nomAttributs.size();
            
            //Iteration sur les noms de propriete
            for( long p = 0; p < tp; p++ )
            {
                //Completer les en-têtes dans le fichier excel
                ws.cell( ( 16 + tt + ( p + 1 ) ), 1 ).value( nomPropriete[p] );
            }
            
            //Barre de progression
            ProgressBar prog_xls = ProgressBar( _T( "Progression :" ), size_blocks );
            
            //Indice de ligne
            int ro = 2;
            
            //Iterer sur les blockes
            for( long i = 0; i < v_blocks.size(); i++ )
            {
                //Recuperer le i-eme block
                blk = v_blocks[i];
                
                //Recuperer les infos poly
                for( int k = 0; k < v_poly2d.size(); k++ )
                {
                    //Recuperer le k-ième polyline 2d
                    poly_2d = v_poly2d[k];
                    
                    //Verifier
                    if( !poly_2d )
                        continue;
                        
                    //Recuperer la position du block courant
                    AcGePoint3d pos_3d = blk->position();;
                    
                    //Transformer le point 3d en 2d pour pouvoir s'en servir dans la fonction isPointInPoly(poly2d, pt2d)
                    AcGePoint2d pos = AcGePoint2d::kOrigin;
                    pos.x = pos_3d.x;
                    pos.y = pos_3d.y;
                    
                    //Booleen pour exprimer le resultat
                    bool inPoly = isPointInPoly( poly_2d, pos );
                    
                    //Si le block n'est pas dans la polyline, continuer
                    if( !inPoly )
                    {
                        poly_2d->close();
                        continue;
                    }
                    
                    //Sinon stocker ce polyline dans vecPoly2d
                    vecPoly2d.emplace_back( poly_2d );
                    
                    // Fermer la polyline
                    poly_2d->close();
                }
                
                //Si le block n'est pas entouré de/des polyline(s) : on continue avec un autre block
                if( vecPoly2d.size() == 0 )
                {
                    //  Incrémentation du nombre de block ignoré
                    ignoredBlock += 1;
                    
                    //  Ajouter le block ignoré dans la séléction
                    acdbGetAdsName( ssSelection, blk->objectId() );
                    acedSSAdd( ssSelection, ssEnd, ssEnd );
                    acedSSFree( ssSelection );
                    
                    //Fermeture du block
                    blk->close();
                    continue;
                }
                
                //Declarer un identifiant unique
                AcDbHandle poly_handle;
                
                //Longueur de la polyline
                double poly_length;
                
                //Surface de la polyline
                double poly_area;
                
                //Surface test
                double area = 1000000;
                
                //Indice du polyline qui a la plus petite surface dans vecPoly2D
                int pos_vp = 0;
                
                //Sinon : on cherche la polyline la plus proche et ayant la plus petite surface
                if( vecPoly2d.size() == 1 )
                {
                    //S'il n'y a qu'une polyline : recuperer son handle
                    poly_handle = vecPoly2d[0]->handle();
                    
                    //Recuperer la longueur de la polyline
                    poly_length = getLength( vecPoly2d[0] );
                    
                    //Recuperer la surface de la polyline
                    vecPoly2d[0]->getArea( poly_area );
                }
                
                else if( vecPoly2d.size() > 1 )
                {
                    //On parcours les polylines et on prend celle qui a la plus petite surface
                    for( int vp = 0; vp < vecPoly2d.size(); vp++ )
                    {
                        //Stocker la surface de la polyline courante
                        double temp_area = 0;
                        vecPoly2d[vp]->getArea( temp_area );
                        
                        //Tester si elle est petite que surface test
                        if( temp_area <= area )
                        {
                            //Recuperer sa surface
                            area = temp_area;
                            
                            //Recuperer son indice
                            pos_vp = vp;
                        }
                        
                        //Fermer la polyline
                        vecPoly2d[vp]->close();
                    }
                    
                    //Recuperer l'identificateur du polyline qui a la plus petite surface
                    poly_handle = vecPoly2d[pos_vp]->handle();
                    
                    //Recuperer la longueur de la polyline
                    poly_length = getLength( vecPoly2d[pos_vp] );
                    
                    //Recuperer la surface de la polyline
                    vecPoly2d[pos_vp]->getArea( poly_area );
                }
                
                // Definir les formats de toutes les cellules dans les colonnes en texte
                ws.columns( true ).number_format( xlnt::number_format::text() );
                
                //Recuperer les données
                exportBlock( blk, pData );
                
                //Ecrire les données
                for( int d = 0; d < pData.size(); d++ )
                {
                    //Convertir les encodages de strings ent utf-8
                    auto data = latin1_to_utf8( acStrToStr( pData[d] ) );
                    
                    //Inserer les informations
                    ws.cell( d + 1, ro ).value( data );
                }
                
                //Inserer l'identifiant unique de la polyline
                ws.cell( 14, ro ).value( poly_handle.ascii() );
                
                //Inserer la longueur de la polyline
                ws.cell( 15, ro ).value( to_string( poly_length ) );
                
                //Inserer la surface de la polyline
                ws.cell( 16, ro ).value( to_string( poly_area ) );
                
                //Iteration sur les attributs
                for( long att = 0; att < nomAttributs.size(); att++ )
                {
                    if( getAttributValue( blk, nomAttributs[att] ) != "empty" )
                        ws.cell( 16 + ( att + 1 ), ro ).value( getAttributValue( blk, nomAttributs[att] ) );
                    else
                        ws.cell( 16 + ( att + 1 ), ro ).value( "N/A" );
                }
                
                // Check si c'est un bloc dynamique
                if( isDynamicBlock( blk ) )
                {
                    for( long propName = 0; propName < nomPropriete.size(); propName++ )
                    {
                        /*
                            OBTENIR LES VALEURS DES PROPRIETES A L'AIDE DES NOMS DE PROPRIETES
                        */
                        
                        AcString nProp = getPropertyStrFromBlock( blk, nomPropriete[propName] );
                        
                        // Si la valeur du propriete est un AcString
                        if( nProp != _T( "" ) )
                        {
                            //Recuperer la taille de nomAttributs
                            int tai_n = nomAttributs.size();
                            
                            //Recuperer le nom de propriete
                            auto pro_name = acStrToStr( getPropertyStrFromBlock( blk, nomPropriete[propName] ) );
                            
                            // Completer les colonnes excel avec les valeurs de proprietes
                            ws.cell( ( 14 + tai_n + 1 + propName ), ro ).value( pro_name );
                        }
                        
                        // Si la valeur du propriete n'est pas un AcString
                        else
                        {
                            // Si la valeur est un double
                            if( getPropertyFromBlock( blk, nomPropriete[propName] ) != -1 || getPropertyFromBlock( blk, nomPropriete[propName] ) != 0 )
                            {
                                // Ecrire les valeurs des proprietes dans les colonnes correspondantes
                                ws.cell( ( 16 + nomAttributs.size() + 1 + propName ), ro ).value( to_string( getPropertyFromBlock( blk, nomPropriete[propName] ) ) );
                            }
                            
                            // Si ce n'est pas dans les deux cas
                            else
                            {
                                // Mettre N/A à la place
                                ws.cell( ( 16 + nomAttributs.size() + 1 + propName ), ro ).value( "N/A" );
                            }
                        }
                    }
                }
                
                // Si ce n'est pas un blooc dynamique
                else
                {
                    // Iteration sur les noms de proprietes
                    for( long pNm = 0; pNm < nomPropriete.size(); pNm++ )
                        // Completer par N/A
                        ws.cell( ( 16 + nomAttributs.size() + 1 + pNm ), ro ).value( "N/A" );
                }
                
                //Incrementation de obj
                obj++;
                
                //Vider vecPoly
                vecPoly2d.clear();
                
                //Fermeture du block
                blk->close();
                
                //Incrementer la ligne
                ro++;
                
                // Vider le tableau des données
                pData.clear();
                
                //Progresser
                prog_xls.moveUp( i );
            }
            
            try
            {
                //Sauvegarder le fichier excel
                wb.save( filename );
                
                //Affichage dans la console le nombre de bloc(s) ou on a exporter ses informations + le nombre de bloc(s) dynamique dans la selection
                print( "Exportation de " + to_string( obj ) + " block(s) dont " + to_string( dynBlocCount ) + " dynamique(s) dans " + acStrToStr( file ) + " terminée." );
                
                //  Affichage dans la console le nombre de block(s) ignoré(s)
                print( "Nombre de block ignoré = " + to_string( ignoredBlock ) );
                
                //Réinitialisation de obj
                obj = 0;
            }
            
            catch( ... )
            {
                print( "Impossible d'enregistrer le fichier." );
                acedSSFree( ssBlock );
                acedSSFree( ssPoly2d );
                return;
            }
            
        }
        
        //Si c'est un fichier .shp
        else if( ext == "shp" )
        {
            //Lire le fichier de parametre
            vector<AcString> field = createBlockPoly2dField();
            writePrjOrNot( file );
            
            //Recuperer les points du surfaces
            vector<AcGePoint3d> point;
            map<int, vector<AcGePoint3d>> mapPoint;
            
            //Creer un fichier shp
            SHPHandle myShapeFile = SHPCreate( file, SHPT_POLYGONZ );
            
            //Creer un dbfFile
            DBFHandle dbfHandle = DBFCreate( file );
            
            //Creer les champs
            int fieldNumber = createFields( dbfHandle, field );
            
            // Initialistion de la barre de progression
            ProgressBar prog = ProgressBar( _T( "Progression" ), size_blocks );
            
            double* x;    double* y;    double* z;
            int exported = 0;
            
            //  Itération sur les blocks
            for( int i = 0; i < v_blocks.size(); i++ )
            {
                //  Entier de vérification du block ignoré (si = 0 : le block est ignoré)
                bool isIgnored = false;
                
                //Recuperer les blocks
                AcDbBlockReference* block = v_blocks[i];
                
                //Verifier block
                if( !block )
                    continue;
                    
                //surface index
                int iS; bool toDraw = false;
                int erea = INT32_MAX;
                
                //Recuperer son point d'insertion
                AcGePoint3d ptBlock = block->position();
                
                //Recuperer l'object id de la polyline
                AcDbObjectId idPoly;
                
                //Boucler sur le polyligne
                for( int j = 0; j < v_poly2d.size(); j++ )
                {
                    //Recuperer le polyligne
                    AcDbPolyline* poly = v_poly2d[j];
                    
                    //Verifier poly
                    if( !poly )
                        continue;
                        
                    //Recuperer le surface
                    double surface;
                    poly->getArea( surface );
                    surface = abs( surface );
                    
                    //Verifier si le block est a l'interieur du polyligne
                    if( isPointInPoly( poly, AcGePoint2d( ptBlock.x, ptBlock.y ) ) )
                    {
                        if( surface < erea )
                        {
                            erea = surface;
                            toDraw = true;
                            iS = j;
                            
                            if( mapPoint.find( j ) == mapPoint.end() )
                            {
                                //Recuperer les vertexs
                                int vertNumber = poly->numVerts();
                                
                                point.resize( 0 );
                                AcGePoint3d p;
                                
                                //Boucler sur les vecteurs
                                for( int k = 0; k < vertNumber; k++ )
                                {
                                    poly->getPointAt( k, p );
                                    point.push_back( p );
                                }
                                
                                //Recuperer le point sur le sommet numero
                                poly->getPointAt( 0, p );
                                point.push_back( p );
                                
                                //Recuperer l'object id de la polyline
                                idPoly = poly->id();
                                
                                //Ajoute les infos dans la map
                                mapPoint.insert( { j, point } );
                            }
                        }
                        
                        //  Mettre le booléen à false
                        isIgnored = false;
                        
                        //  Fermer la polyligne
                        poly->close();
                        
                        //  Terminer le boucle sur les polylignes
                        break;
                    }
                    
                    //  Toujours garder le booléen à true (block ignoré)
                    isIgnored = true;
                    
                    //  Fermer la polyligne
                    poly->close();
                }
                
                AcDbEntity* entPoly = NULL;
                
                if( toDraw )
                {
                    acdbOpenAcDbEntity( entPoly, idPoly, AcDb::kForRead );
                    
                    if( !entPoly )
                        continue;
                        
                    //Recuperer le vecteur point du rang du sommets
                    vector<AcGePoint3d> vecPoint = mapPoint.at( iS );
                    drawBlockPoly( block, entPoly, vecPoint, myShapeFile, dbfHandle, field );
                    exported++;
                }
                
                if( isIgnored )
                {
                    //  Ajouter le block ignoré dans la séléction
                    acdbGetAdsName( ssSelection, block->objectId() );
                    acedSSAdd( ssSelection, ssEnd, ssEnd );
                    acedSSFree( ssSelection );
                    
                    //  Incrémentation du nombre de block ignoré
                    ignoredBlock += 1;
                }
                
                prog.moveUp( i );
                //Fermer le block
                block->close();
                
                // incrementer la barre de progression
                prog.moveUp( i );
            }
            
            DBFClose( dbfHandle );
            SHPClose( myShapeFile );
            
            //Message
            print( "Exportation de " + to_string( exported ) + " block(s) dans " + acStrToStr( file ) + " terminée." );
            
            //  Affichage dans la console le nombre de block ignoré
            print( "Nombre de block ignoré = " + to_string( ignoredBlock ) );
        }
    }
    
    // Si aucun bloc n'est selectionné
    else
        print( "Aucun block sélectionné" );
        
    //  Ajouter les blocks ignorés dans la séléction
    acedSSSetFirst( ssEnd, ssEnd );
    
    //Liberer la mémoire
    acedSSFree( ssBlock );
    acedSSFree( ssPoly2d );
}

void cmdExportBlockPoly3D()
{
    //Définition de la selection
    ads_name ssBlock;
    ads_name ssPoly3d;
    ads_name ssError, sserror;
    
    //Vecteur des blocks ignorés
    vector<AcDbObjectId> vec_error_blk;
    
    //Fichier & Extension
    AcString file, ext;
    
    file = _T( "" );
    
    //Recuperer le  chemin du dwg courant
    AcString current_folder = getCurrentFileFolder();
    current_folder += _T( "\\" );
    
    //Stockage : propriétés & données
    AcStringArray pData;
    AcStringArray propName;
    
    //Vecteur temporaire: liste des attributs & liste des noms de propriete
    vector<AcString> nomAttributs, nomPropriete;
    
    //Vecteur de type non fixe pour contenir les valeurs des proprietes
    vector<void*> propValue;
    
    //Set : (Noms des attributs sans doublons & Noms des propriétes sans doublons & Noms de définition) pour toutes les entités
    set<AcString> sAttr, sProp;
    
    //Nombre de blocs dynamiques
    int dynBlocCount = 0;
    
    //Message
    print( "Veuillez sélectionner le/les block(s) : " );
    
    //Recuperer le nombre de blockes
    long size = getSsBlock( ssBlock, "", "" );
    
    //Declarer un vecteur de polyline pour stocker les polylines entourant un block
    vector<AcDb3dPolyline*> vecPoly3d;
    
    //Déclaration d'un AcDbPolyline
    AcDb3dPolyline* poly_3d = NULL;
    
    //Message
    print( "Veuillez sélectionner la/les polyline(s) : " );
    
    //Selectionner tous les polylines fermées et recuperes leur nombre
    long size_p3d = getSsPoly3D( ssPoly3d, _T( "" ), true );
    
    //Verifier
    if( size != 0 )
    {
        //Nombre d'objet exporté
        int obj = 0;
        
        //Demander chemin du fichier vers quoi on exporte
        file = askForFilePath( false,
                "xlsx;shp;csv;txt",
                "Enregistrer sous",
                current_folder );
                
        //Gestion d'erreur
        if( file == _T( "" ) )
        {
            print( "Commande annulée." );
            acedSSFree( ssBlock );
            acedSSFree( ssPoly3d );
            return;
        }
        
        //Recuperer l'extension
        ext = getFileExt( file );
        
        //Changer l'encodage du chemin vers le fichier excel en utf8
        auto filename = latin1_to_utf8( acStrToStr( file ) );
        
        //Déclarer une reference de blocke
        AcDbBlockReference* blk;
        
        //Fichier : .txt ou .csv
        if( ext == "txt" || ext == "csv" )
        {
            //Ouvrir le fichier
            fstream fichier( acStrToStr( file ), ios::app );
            
            //Ecriture des en-têtes du fichier
            fichier << "Handle;Calque;Nom;Couleur;Type de ligne;Echelle de type de ligne;Epaisseur de ligne;Transparence;Materiau;X;Y;Z;Rotation;Handle poly;Longueur 2D;Longueur 3D; Surface";
            
            //Itération sur les objets
            for( long i = 0; i < size; i++ )
            {
                //Pointeur sur le bloc
                blk = getBlockFromSs( ssBlock, i );
                
                //Verifier
                if( !blk )
                    continue;
                    
                //Recuperer la taille
                int taille_attr = getAttributesNamesListOfBlockRef( blk ).size();
                
                //Iteration sur la taille de la liste des noms des attributs
                for( long j = 0; j < taille_attr; j++ )
                {
                    //Mettre les attributs dans le set
                    sAttr.insert( getAttributesNamesListOfBlockRef( blk )[j] );
                }
                
                //Si block dynamique
                if( isDynamicBlock( blk ) )
                {
                    //Inrementation du nombre de bloc dynamique
                    dynBlocCount++;
                    
                    //Prendre les listes des noms de proprietes et aussi des valeurs de proprietes
                    getBlockPropWithValueList( blk, propName, propValue );
                    
                    //Iteration sur la liste des noms des proprietes
                    for( long pn = 0; pn < propName.size(); pn++ )
                    {
                        //Mettre les noms des proprietes dans le set (sans doublons)
                        sProp.insert( propName[pn] );
                    }
                }
                
                //Fermeture du bloc
                blk->close();
            }
            
            //Itération sur les données sans doublons dans le |setAttr| et les mettre dans le vecteur temporaire
            for( set<AcString>::iterator t = sAttr.begin(); t != sAttr.end(); t++ )
            {
                nomAttributs.push_back( *t );
                //Completer les en-tête avec la liste sans doublons des ATTRIBUTS
                fichier << ";" << acStrToStr( *t );
            }
            
            //Iteration sur les données sans doublons dans le |sProp| et les mettre dans le vecteur temporaire
            for( set<AcString>::iterator r = sProp.begin(); r != sProp.end(); r++ )
            {
                //Mettre les noms de proprietés sans doublons dans un vecteur
                nomPropriete.push_back( *r );
                //Completer les en-têtes avec la liste sans doublons des PROPRIETES
                fichier << ";" << acStrToStr( *r );
            }
            
            //Se mettre à la ligne pour commencer à inserer les données
            fichier << endl;
            
            //Barre de progression
            ProgressBar prog = ProgressBar( _T( "Progression :" ), size );
            
            //Reiterer sur les blocks
            for( long i = 0; i < size; i++ )
            {
            
                //Pointeur sur le block
                blk = getBlockFromSs( ssBlock, i );
                
                //Verifier blk
                if( !blk )
                    continue;
                    
                //Recuperer les infos poly
                for( int k = 0; k < size_p3d; k++ )
                {
                    //Recuperer le k-ième polyline 3d
                    poly_3d = getPoly3DFromSs( ssPoly3d, k );
                    
                    //Verifier
                    if( !poly_3d )
                        continue;
                        
                    //Recuperer la position du block courant
                    AcGePoint3d pos_3d = blk->position();;
                    
                    
                    //Booleen pour exprimer le resultat
                    bool inPoly = isPointInPoly( poly_3d, pos_3d );
                    
                    //Si le block n'est pas dans la polyline, continuer
                    if( !inPoly )
                    {
                        poly_3d->close();
                        continue;
                    }
                    
                    //Sinon stocker ce polyline dans vecPoly3d
                    vecPoly3d.emplace_back( poly_3d );
                    
                    // Fermer la polyline
                    poly_3d->close();
                }
                
                //Si le block n'est pas entouré de/des polyline(s) : on continue avec un autre block
                if( vecPoly3d.size() == 0 )
                {
                    //Inserer le block dans vec_error_blk
                    vec_error_blk.emplace_back( blk->id() );
                    
                    //Fermeture du block
                    blk->close();
                    continue;
                }
                
                //Declarer un identifiant unique
                AcDbHandle poly_handle;
                
                //Declarer la longueur de la polyline
                double poly_length;
                
                //Declarer la longueur de la polyline 3d
                double poly_length3d;
                
                //Declarer la surface de la polyline
                double poly_area;
                
                //Surface test
                double area = 1000000;
                
                //Indice du polyline qui a la plus petite surface dans vecPoly3D
                int pos_vp = 0;
                
                //Sinon : on cherche la polyline la plus proche et ayant la plus petite surface
                if( vecPoly3d.size() == 1 )
                {
                    //S'il n'y a qu'une polyline : recuperer son handle
                    poly_handle = vecPoly3d[0]->handle();
                    
                    //Recuperer la longueur 2d de la polyline
                    poly_length = get2DLengthPoly( vecPoly3d[0] );
                    
                    //Recuperer la longueur 3d de la polyline
                    poly_length3d = getLength( vecPoly3d[0] );
                    
                    //Recuperer la surface
                    vecPoly3d[0]->getArea( poly_area );
                    
                    vecPoly3d[0]->close();
                }
                
                else if( vecPoly3d.size() > 1 )
                {
                    //On parcours les polylines et on prend celle qui a la plus petite surface
                    for( int vp = 0; vp < vecPoly3d.size(); vp++ )
                    {
                        //Stocker la surface de la polyline courante
                        double temp_area = 0;
                        vecPoly3d[vp]->getArea( temp_area );
                        
                        //Tester si elle est petite que surface test
                        if( temp_area <= area )
                        {
                            //Recuperer sa surface
                            area = temp_area;
                            
                            //Recuperer son indice
                            pos_vp = vp;
                        }
                        
                        //Fermer la polyline
                        vecPoly3d[vp]->close();
                    }
                    
                    //Recuperer l'identificateur du polyline qui a la plus petite surface
                    poly_handle = vecPoly3d[pos_vp]->handle();
                    
                    //Recuperer la longueur 2d de la polyline
                    poly_length = get2DLengthPoly( vecPoly3d[pos_vp] );
                    
                    //Recuperer la longueur 3d de la polyline
                    poly_length3d = getLength( vecPoly3d[pos_vp] );
                    
                    //Recuperer la surface
                    vecPoly3d[pos_vp]->getArea( poly_area );
                    
                    vecPoly3d[pos_vp]->close();
                }
                
                // Appel de la fonction de traitement
                exportBlock( blk, pData );
                
                // Iteration sur les données
                for( long j = 0; j < pData.size(); j++ )
                {
                    //Ecrire les données
                    fichier << pData[j] << ";";
                }
                
                //Inserer l'identifiant unique de la polyline
                fichier << poly_handle.ascii() << ";";
                
                //Inserer la longueur 2d de la polyline
                fichier << poly_length << ";";
                
                //Inserer la longueur 3d de la polyline
                fichier << poly_length3d << ";";
                
                //Inserer la surface de la polyline
                fichier << poly_area << ";";
                
                //Rcuperer la taille de nomAttributs
                int nb_att = nomAttributs.size();
                
                //Itération sur les attributs
                for( long att = 0; att < nb_att; att++ )
                {
                    //Si la valeur de l'attribut n'est pas vide
                    if( getAttributValue( blk, nomAttributs[att] ) != "empty" )
                    {
                        //Convertir en std::string
                        auto attrib_val = acStrToStr( getAttributValue( blk, nomAttributs[att] ) );
                        
                        //Ecrire la valeur de l'attribut
                        fichier << attrib_val << ";";
                    }
                    
                    //Si la valeur de l'attribut est vide
                    else
                    {
                        // Remplir la colonne par des N/A
                        fichier << "N/A;";
                    }
                }
                
                //Check si c'est un bloc dynamique
                if( isDynamicBlock( blk ) )
                {
                
                    //Iteration sur les noms de proprieté sans doublons
                    for( long propName = 0; propName < nomPropriete.size(); propName++ )
                    {
                        //Recuperer la valeur du propriete
                        AcString bProp = getPropertyStrFromBlock( blk, nomPropriete[propName] );
                        
                        //Si bProp n'est pas vide
                        if( bProp != _T( "" ) )
                        {
                            //Recuperer la valeur du propriete
                            auto prop_val = acStrToStr( getPropertyStrFromBlock( blk, nomPropriete[propName] ) );
                            
                            //Ecrire dans le fichier
                            fichier << prop_val << ";";
                        }
                        
                        // Si la valeur du propriete n'est pas vide
                        else
                        {
                            // Si la valeur du propriete est un double
                            if( getPropertyFromBlock( blk, nomPropriete[propName] ) != -1 || getPropertyFromBlock( blk, nomPropriete[propName] ) != 0 )
                            {
                                //Recuperer la propriete
                                auto prop_value = to_string( getPropertyFromBlock( blk, nomPropriete[propName] ) );
                                
                                //Ecrire les valeurs des proprietés dans le fichier
                                fichier << prop_value << ";";
                            }
                            
                            // Si ce n'est pas les deux types
                            else
                                fichier << "N/A;";
                        }
                        
                    }
                    
                }
                
                // Si ce n'est pas un bloc dynamique
                else
                {
                    for( long propName = 0; propName < nomPropriete.size(); propName++ )
                        fichier << "N/A;";
                }
                
                // Se mettre à la ligne après une écriture d'une ligne d'information des objets
                fichier << endl;
                
                //Incrémentation de obj
                obj++;
                
                //Fermeture du block
                blk->close();
                
                // Vider le tableau des données
                pData.clear();
                
                //Vider vecPoly
                vecPoly3d.clear();
                
                //Progresser
                prog.moveUp( i );
            }
            
            //Recuperer le nombre de blocks ignorés
            int taille_blk_ignored = vec_error_blk.size();
            
            if( taille_blk_ignored > 0 )
            {
                for( auto& obj_id : vec_error_blk )
                {
                    acdbGetAdsName( ssError, obj_id );
                    acedSSAdd( ssError, sserror, sserror );
                    acedSSFree( ssError );
                }
            }
            
            try
            {
                //Fermeture du fichier
                fichier.close();
                
                //Affichage dans la console que les informations sont bien exportés
                print( "Exportation de " + to_string( obj ) + " block(s) dont " + to_string( dynBlocCount ) + " bloc(s) dynamique(s) dans " + acStrToStr( file ) + " terminée." );
                print( "Nombre de blocks ignorés : " + to_string( taille_blk_ignored ) );
                //Réinitialisation de obj
                obj = 0;
            }
            
            catch( ... )
            {
                print( "Impossible d'enregistrer le fichier." );
                acedSSFree( ssBlock );
                acedSSFree( ssPoly3d );
                return;
            }
            
        }
        
        // Si c'est un fichier .xlsx
        else if( ext == "xlsx" )
        {
        
            //Initialisation du fichier excel
            workbook wb;
            worksheet ws = wb.active_sheet();
            
            //Ajout des en-tête
            ws.cell( 1, 1 ).value( "Handle" );
            ws.cell( 2, 1 ).value( "Calque" );
            ws.cell( 3, 1 ).value( "Nom" );
            ws.cell( 4, 1 ).value( "Couleur" );
            ws.cell( 5, 1 ).value( "Type de ligne" );
            ws.cell( 6, 1 ).value( "Echelle de type de ligne" );
            ws.cell( 7, 1 ).value( "Epaisseur de ligne" );
            ws.cell( 8, 1 ).value( "Transparence" );
            ws.cell( 9, 1 ).value( "Materiau" );
            ws.cell( 10, 1 ).value( "X" );
            ws.cell( 11, 1 ).value( "Y" );
            ws.cell( 12, 1 ).value( "Z" );
            ws.cell( 13, 1 ).value( "Rotation" );
            ws.cell( 14, 1 ).value( "Handle poly" );
            ws.cell( 15, 1 ).value( "Longueur 2d" );
            ws.cell( 16, 1 ).value( "Longueur 3d" );
            ws.cell( 17, 1 ).value( "Surface" );
            
            //Itération sur les objets
            for( long i = 0; i < size; i++ )
            {
            
                //Pointeur sur le bloc
                blk = getBlockFromSs( ssBlock, i );
                
                //Verifier le blk
                if( !blk )
                    continue;
                    
                //Recuperer la taille
                int taille_attr = getAttributesNamesListOfBlockRef( blk ).size();
                
                //Iteration sur la taille de la liste des noms des attributs
                for( long j = 0; j < taille_attr; j++ )
                {
                    //Mettre les attributs dans le set
                    sAttr.insert( getAttributesNamesListOfBlockRef( blk )[j] );
                }
                
                //Si block dynamique
                if( isDynamicBlock( blk ) )
                {
                    //Inrementation du nombre de bloc dynamique
                    dynBlocCount++;
                    
                    //Prendre les listes des noms de proprietes et aussi des valeurs de proprietes
                    getBlockPropWithValueList( blk, propName, propValue );
                    
                    //Iteration sur la liste des noms des proprietes
                    for( long pn = 0; pn < propName.size(); pn++ )
                    {
                        //Mettre les noms des proprietes dans le set (sans doublons)
                        sProp.insert( propName[pn] );
                    }
                }
                
                //Fermeture du bloc
                blk->close();
            }
            
            //Iteration sur les données sans doublons dans le |setAttr| et les mettre dans le vecteur
            for( set<AcString>::iterator t = sAttr.begin(); t != sAttr.end(); t++ )
            {
                // Mettre les données sans doublons dans un vecteur
                nomAttributs.push_back( *t );
            }
            
            //Recuperer la taille de nomAttributs
            int ta = nomAttributs.size();
            
            //Completer les colonnes d'en tête
            for( int a = 0; a < ta; a++ )
            {
                // Completer les données dans le fichier excel
                ws.cell( 17 + ( a + 1 ), 1 ).value( nomAttributs[a] );
            }
            
            
            //Iteration sur les données sans doublons dans le |sProp| et les mettre dans le vecteur temporaire
            for( set<AcString>::iterator r = sProp.begin(); r != sProp.end(); r++ )
            {
                //Mettre les noms des proprietes sans doublons dans le vecteur
                nomPropriete.push_back( *r );
            }
            
            //Recuperer la taille du nomPropriete
            int tp = nomPropriete.size();
            
            //Recuperer la taille du nomAttributs
            int tt = nomAttributs.size();
            
            //Iteration sur les noms de propriete
            for( long p = 0; p < tp; p++ )
            {
                //Completer les en-têtes dans le fichier excel
                ws.cell( ( 17 + tt + ( p + 1 ) ), 1 ).value( nomPropriete[p] );
            }
            
            //Barre de progression
            ProgressBar prog_xls = ProgressBar( _T( "Progression :" ), size );
            
            //Indice de ligne
            int ro = 2;
            
            //Iterer sur les blockes
            for( long i = 0; i < size; i++ )
            {
            
                //Recuperer le i-eme block
                blk = getBlockFromSs( ssBlock, i, GcDb::kForRead );
                
                //Recuperer les infos poly
                for( int k = 0; k < size_p3d; k++ )
                {
                    //Recuperer le k-ième polyline 3d
                    poly_3d = getPoly3DFromSs( ssPoly3d, k );
                    
                    //Verifier
                    if( !poly_3d )
                        continue;
                        
                    //Recuperer la position du block courant
                    AcGePoint3d pos_3d = blk->position();;
                    
                    //Booleen pour exprimer le resultat
                    bool inPoly = isPointInPoly( poly_3d, pos_3d );
                    
                    //Si le block n'est pas dans la polyline, continuer
                    if( !inPoly )
                    {
                        poly_3d->close();
                        continue;
                    }
                    
                    //Sinon stocker ce polyline dans vecPoly2d
                    vecPoly3d.emplace_back( poly_3d );
                    
                    // Fermer la polyline
                    poly_3d->close();
                }
                
                //Si le block n'est pas entouré de/des polyline(s) : on continue avec un autre block
                if( vecPoly3d.size() == 0 )
                {
                    //Recuperer les blocks ignorés
                    vec_error_blk.emplace_back( blk->id() );
                    
                    //Fermeture du block
                    blk->close();
                    
                    continue;
                }
                
                //Declarer un identifiant unique
                AcDbHandle poly_handle;
                
                //Declarer la longueur de la polyline
                double poly_length;
                
                //Declarer la longueur de la polyline 3d
                double poly_length3d;
                
                //Declarer la surface de la polyline
                double poly_area;
                
                //Surface test
                double area = 1000000;
                
                //Indice du polyline qui a la plus petite surface dans vecPoly3D
                int pos_vp = 0;
                
                //Sinon : on cherche la polyline la plus proche et ayant la plus petite surface
                if( vecPoly3d.size() == 1 )
                {
                    //S'il n'y a qu'une polyline : recuperer son handle
                    poly_handle = vecPoly3d[0]->handle();
                    
                    //Recuperer la longueur 2d de la polyline
                    poly_length = get2DLengthPoly( vecPoly3d[0] );
                    
                    //Recuperer la longueur 3d de la polyline
                    poly_length3d = getLength( vecPoly3d[0] );
                    
                    //Recuperer la surface
                    vecPoly3d[0]->getArea( poly_area );
                    
                    //Liberer la mémoire
                    vecPoly3d[0]->close();
                }
                
                else if( vecPoly3d.size() > 1 )
                {
                    //On parcours les polylines et on prend celle qui a la plus petite surface
                    for( int vp = 0; vp < vecPoly3d.size(); vp++ )
                    {
                        //Stocker la surface de la polyline courante
                        double temp_area = 0;
                        vecPoly3d[vp]->getArea( temp_area );
                        
                        //Tester si elle est petite que surface test
                        if( temp_area <= area )
                        {
                            //Recuperer sa surface
                            area = temp_area;
                            
                            //Recuperer son indice
                            pos_vp = vp;
                        }
                        
                        //Fermer la polyline
                        vecPoly3d[vp]->close();
                    }
                    
                    //Recuperer l'identificateur du polyline qui a la plus petite surface
                    poly_handle = vecPoly3d[pos_vp]->handle();
                    
                    //Recuperer la longueur 2d de la polyline
                    poly_length = get2DLengthPoly( vecPoly3d[pos_vp] );
                    
                    //Recuperer la longueur 3d de la polyline
                    poly_length3d = getLength( vecPoly3d[pos_vp] );
                    
                    //Recuperer la surface
                    vecPoly3d[pos_vp]->getArea( poly_area );
                    
                    //Liberer la mémoire
                    vecPoly3d[pos_vp]->close();
                }
                
                // Definir les formats de toutes les cellules dans les colonnes en texte
                ws.columns( true ).number_format( xlnt::number_format::text() );
                
                //Recuperer les données
                exportBlock( blk, pData );
                
                
                //Ecrire les données
                for( int d = 0; d < pData.size(); d++ )
                {
                    //Convertir les encodages de strings ent utf-8
                    auto data = latin1_to_utf8( acStrToStr( pData[d] ) );
                    
                    ws.cell( d + 1, ro ).value( data );
                }
                
                //Inserer l'identifiant unique de la polyline
                ws.cell( 14, ro ).value( poly_handle.ascii() );
                
                //Inserer la longueur 2d de la polyline
                ws.cell( 15, ro ).value( to_string( poly_length ) );
                
                //Inserer la longueur 3d de la polyline
                ws.cell( 16, ro ).value( to_string( poly_length3d ) );
                
                //Inserer la surface de la polyline
                ws.cell( 17, ro ).value( to_string( poly_area ) );
                
                //Iteration sur les attributs
                for( long att = 0; att < nomAttributs.size(); att++ )
                {
                    if( getAttributValue( blk, nomAttributs[att] ) != "empty" )
                        ws.cell( 17 + ( att + 1 ), ro ).value( getAttributValue( blk, nomAttributs[att] ) );
                    else
                        ws.cell( 17 + ( att + 1 ), ro ).value( "N/A" );
                }
                
                // Check si c'est un bloc dynamique
                if( isDynamicBlock( blk ) )
                {
                    for( long propName = 0; propName < nomPropriete.size(); propName++ )
                    {
                        AcString nProp = getPropertyStrFromBlock( blk, nomPropriete[propName] );
                        
                        // Si la valeur du propriete est un AcString
                        if( nProp != _T( "" ) )
                        {
                            //Recuperer la taille de nomAttributs
                            int tai_n = nomAttributs.size();
                            
                            //Recuperer le nom de propriete
                            auto pro_name = acStrToStr( getPropertyStrFromBlock( blk, nomPropriete[propName] ) );
                            
                            // Completer les colonnes excel avec les valeurs de proprietes
                            ws.cell( ( 17 + tai_n + 1 + propName ), ro ).value( pro_name );
                        }
                        
                        // Si la valeur du propriete n'est pas un AcString
                        else
                        {
                            // Si la valeur est un double
                            if( getPropertyFromBlock( blk, nomPropriete[propName] ) != -1 || getPropertyFromBlock( blk, nomPropriete[propName] ) != 0 )
                            {
                                // Ecrire les valeurs des proprietes dans les colonnes correspondantes
                                ws.cell( ( 17 + nomAttributs.size() + 1 + propName ), ro ).value( to_string( getPropertyFromBlock( blk, nomPropriete[propName] ) ) );
                            }
                            
                            // Si ce n'est pas dans les deux cas
                            else
                            {
                                // Mettre N/A à la place
                                ws.cell( ( 17 + nomAttributs.size() + 1 + propName ), ro ).value( "N/A" );
                            }
                        }
                    }
                }
                
                // Si ce n'est pas un blooc dynamique
                else
                {
                    // Iteration sur les noms de proprietes
                    for( long pNm = 0; pNm < nomPropriete.size(); pNm++ )
                        // Completer par N/A
                        ws.cell( ( 17 + nomAttributs.size() + 1 + pNm ), ro ).value( "N/A" );
                }
                
                //Incrementation de obj
                obj++;
                
                //Vider vecPoly
                vecPoly3d.clear();
                
                //Fermeture du block
                blk->close();
                
                //Incrementer la ligne
                ro++;
                
                // Vider le tableau des données
                pData.clear();
                
                //Progresser
                prog_xls.moveUp( i );
            }
            
            //Recuperer le nombre de blocks ignorés
            int taille_blk_ignored = vec_error_blk.size();
            
            //Ajouter les blocks ignorés dans la selection
            if( taille_blk_ignored > 0 )
            {
                for( auto& obj_id : vec_error_blk )
                {
                    acdbGetAdsName( ssError, obj_id );
                    acedSSAdd( ssError, sserror, sserror );
                    acedSSFree( ssError );
                }
            }
            
            try
            {
                //Sauvegarder le fichier excel
                wb.save( filename );
                
                //Affichage dans la console le nombre de bloc(s) ou on a exporter ses informations + le nombre de bloc(s) dynamique dans la selection
                print( "Exportation de " + to_string( obj ) + " block(s) dont " + to_string( dynBlocCount ) + " dynamique(s) dans " + acStrToStr( file ) + " terminée." );
                print( "Nombre de blocks ignorés : " + to_string( taille_blk_ignored ) );
                
                //Réinitialisation de obj
                obj = 0;
            }
            
            catch( ... )
            {
                print( "Impossible d'enregistrer le fichier." );
                acedSSFree( ssBlock );
                acedSSFree( ssPoly3d );
                return;
            }
            
        }
        
        //Fichier: shape
        else if( ext == "shp" )
        {
            //Lire le fichier de parametre
            vector<AcString> field = createBlockPoly3dField();
            writePrjOrNot( file );
            
            //Recuperer les points du surfaces
            vector<AcGePoint3d> point;
            map<int, vector<AcGePoint3d>> mapPoint;
            
            //Creer un fichier shp
            SHPHandle myShapeFile = SHPCreate( file, SHPT_POLYGONZ );
            
            //Creer un dbfFile
            DBFHandle dbfHandle = DBFCreate( file );
            
            //Creer les champs
            int fieldNumber = createFields( dbfHandle, field );
            
            // Initialistion de la barre de progression
            ProgressBar prog = ProgressBar( _T( "Progression" ), size );
            
            //Nombre d'objets exportés
            int exported = 0;
            
            //Nombre de blocks ignorés
            int ignored_blk = 0;
            
            //Boucler sur la selection
            for( int i = 0; i < size; i++ )
            {
                //Recuperer les blocks
                AcDbBlockReference* block = getBlockFromSs( ssBlock, i, AcDb::kForRead );
                
                //Verifier block
                if( !block )
                    continue;
                    
                //surface index
                int iS; bool toDraw = false;
                int erea = INT32_MAX;
                
                //Recuperer son point d'insertion
                AcGePoint3d ptBlock = block->position();
                
                //Recuperer l'object id de la polyline
                AcDbObjectId idPoly = NULL;
                
                //Boucler sur le polyligne
                for( int j = 0; j < size_p3d; j++ )
                {
                    //Recuperer le polyligne
                    AcDb3dPolyline* poly = getPoly3DFromSs( ssPoly3d, j );
                    
                    //Verifier poly
                    if( !poly )
                        continue;
                        
                    //Recuperer le surface
                    double surface;
                    poly->getArea( surface );
                    surface = abs( surface );
                    
                    //Verifier si le block est a l'interieur du polyligne
                    if( isPointInPoly( poly, ptBlock ) )
                    {
                        if( surface < erea )
                        {
                            erea = surface;
                            toDraw = true;
                            iS = j;
                            
                            if( mapPoint.find( j ) == mapPoint.end() )
                            {
                                point.resize( 0 );
                                AcGePoint3d p;
                                //Boucer sur les sommets
                                AcDbObjectIterator* iter = poly->vertexIterator();
                                AcDb3dPolylineVertex* vertex;
                                
                                for( iter->start(); !iter->done(); iter->step() )
                                {
                                    if( !poly->openVertex( vertex, iter->objectId(), AcDb::kForRead ) )
                                    {
                                        p = vertex->position();
                                        point.push_back( p );
                                    }
                                }
                                
                                //Recuperer l'objectid de la polyline
                                idPoly = poly->id();
                                
                                vertex->close();
                                //Recuperer le point sur le sommet numero
                                poly->getStartPoint( p );
                                point.push_back( p );
                                
                                //Ajoute les infos dans la map
                                mapPoint.insert( { j, point } );
                            }
                        }
                    }
                    
                    
                    poly->close();
                }
                
                AcDbEntity* entPoly = NULL;
                
                if( toDraw )
                {
                    acdbOpenAcDbEntity( entPoly, idPoly, AcDb::kForRead );
                    
                    if( !entPoly )
                        continue;
                        
                    //Recuperer le vecteur point du rang du sommets
                    vector<AcGePoint3d> vecPoint = mapPoint.at( iS );
                    drawBlockPoly( block, entPoly, vecPoint, myShapeFile, dbfHandle, field );
                    exported++;
                }
                
                else
                {
                    acdbGetAdsName( ssError, block->id() );
                    acedSSAdd( ssError, sserror, sserror );
                    acedSSFree( ssError );
                    
                    //Incrementer le nombre des blocks ignorés
                    ignored_blk++;
                }
                
                prog.moveUp( i );
                //Fermer le block
                block->close();
                
                // incrementer la barre de progression
                prog.moveUp( i );
                
            }
            
            //Liberer la mémoire
            DBFClose( dbfHandle );
            SHPClose( myShapeFile );
            
            //Affichage dans la console le nombre de bloc(s) ou on a exporter ses informations  dans la selection
            print( "Exportation de " + to_string( exported ) + " block(s) dans " + acStrToStr( file ) + " terminée." );
            print( "Nombre de blocks ignorés : " + to_string( ignored_blk ) );
        }
        
    }
    
    // Si aucun bloc n'est selectionné
    else
        print( "Commande annulée." );
        
    //Ajouter dans la selection courante les blockes ignorés
    acedSSSetFirst( sserror, sserror );
    
    //Liberer la mémoire
    acedSSFree( ssBlock );
    acedSSFree( ssPoly3d );
    
}

void cmdExportTextPoly2D()
{
    //Nom du texte
    ads_name ssText;
    ads_name ssPoly2d;
    
    ads_name sserror, ssError;
    
    //Declarer les text qui ne sont pas traités
    vector<AcDbText*> vec_error_txt;
    
    //Fichier & Extension
    AcString file, ext;
    
    file = _T( "" );
    
    //Recuperer le  chemin du dwg courant
    AcString current_folder = getCurrentFileFolder();
    current_folder += _T( "\\" );
    
    //Tableau de stockage des données
    AcStringArray pData;
    
    //Déclarer un AcDbPolyline
    AcDbPolyline* poly_2d = NULL;
    
    //Qui va servir à contenir le(s) polyline(s) qui entoure(nt) un text
    vector<AcDbPolyline*> vec_poly;
    
    //Message
    print( "Veuillez séléctionner le/les texte(s) : " );
    
    //Recuperer le nombre de texte
    long size = getSsText( ssText, _T( "" ) );
    
    //Message
    print( "Veuillez séléctionner la/les polyline(s) : " );
    
    //Recuperer le nombre de polylines
    long size_2d = getSsPoly2D( ssPoly2d, _T( "" ) );
    
    
    
    //Vérification s'il existe des textes séléctionnés
    if( size != 0 )
    {
    
        //Nombre d'objet sélectionné
        int obj = 0;
        
        //2 - Demander à l'utilisateur le répertoire, le nom et l'extension du fichier ou il va enregistrer dans un dossier
        file = askForFilePath( false,
                "xlsx;shp;csv;txt",
                "Enregistrer sous",
                current_folder );
                
        //Gestion d'erreur
        if( file == _T( "" ) )
        {
            print( "Commande annulée." );
            acedSSFree( ssText );
            acedSSFree( ssPoly2d );
            return;
            
        }
        
        //Recuperer l'extension
        ext = getFileExt( file );
        
        //Changer l'encodage du chemin vers le fichier excel en utf8
        auto filename = latin1_to_utf8( acStrToStr( file ) );
        
        //Fichier .txt, .csv
        if( ext == "csv" || ext == "txt" )
        {
            //Ouvverture du fichier obtenue dans askForFilePath()
            fstream fichier( acStrToStr( file ), ios::app );
            
            //Ecrire les en-têtes du fichier
            fichier << "Handle;Calque;Couleur;Type de ligne;Echelle type de ligne;Epaisseur de ligne;Transparence;Epaisseur;Materiau;X;Y;Z;Rotation;Valeur;Hauteur;Style;Mode verticale;Mode horizontale;Handle poly;Longueur 2D;Surface" << endl;
            
            //L'entité texte
            AcDbText* txt;
            
            ProgressBar prog_txt = ProgressBar( _T( "Progression:" ), size );
            
            //Iterer sur les objet textes
            for( long i = 0; i < size; i++ )
            {
            
                //Recuperer le i-eme texte
                txt = getTextFromSs( ssText, i, GcDb::kForRead );
                
                //Safeguard
                if( !txt )
                    continue;
                    
                //Recuperer la position du texte
                AcGePoint3d pos_text = txt->position();
                
                //Recuperer l'extents du texte
                AcDbExtents ext;
                txt->getGeomExtents( ext );
                //Recuperer la position du texte
                AcGePoint3d al_pt = midPoint3d( ext.minPoint(), ext.maxPoint() );
                
                //Iterer sur la selection de polylines
                for( int pl = 0; pl < size_2d; pl++ )
                {
                
                    //Recuperer le pl-eme polyligne
                    poly_2d = getPoly2DFromSs( ssPoly2d, pl );
                    
                    //Tester si la polyline est fermée ou non
                    if( !isClosed( poly_2d ) )
                        continue;
                        
                    //Tester si pos_2d est dans poly_2d
                    bool test_txt = isPointInPoly( poly_2d, AcGePoint2d( al_pt.x, al_pt.y ) );
                    
                    //Verifier
                    if( !test_txt )
                    {
                        poly_2d->close();
                        continue;
                    }
                    
                    //Inserer la polyline dans vec_poly
                    vec_poly.emplace_back( poly_2d );
                    
                    //Fermer la polyline
                    poly_2d->close();
                }
                
                //Déclarer un identificateur unique
                AcDbHandle poly_handle;
                
                //Longueur 2d de la polyline
                double poly_length;
                
                //Surface de la polyline
                double poly_area;
                
                //Si aucune polyline n'entoure le text, continuer la boucle de text
                if( vec_poly.size() == 0 )
                {
                    //Inserer le texte dans la liste d'erreur
                    vec_error_txt.emplace_back( txt );
                    
                    txt->close();
                    continue;
                }
                
                //Parcourir les polylines dans vec_poly
                if( vec_poly.size() == 1 )
                {
                    //Affecter le handle de la polyline
                    poly_handle = vec_poly[0]->handle();
                    
                    //Recuperer la longueur de la polyline
                    poly_length = getLength( vec_poly[0] );
                    
                    //Recuperer la surface de la polyline
                    vec_poly[0]->getArea( poly_area );
                    
                    //Liberer la selection
                    vec_poly[0]->close();
                }
                
                else
                {
                    //Trouver la polyline qui a la plus petite surface
                    double area_temp = 100000000;
                    
                    //Position de ce polyline dans vec_poly
                    int pos_poly = 0;
                    
                    //Iterer sur vec_poly
                    for( int h = 0; h < vec_poly.size(); h++ )
                    {
                        //temp
                        double area = 0;
                        
                        //Recuperer la surface de la polyline
                        vec_poly[h]->getArea( area );
                        
                        //Recuperer la surface et la position
                        if( area < area_temp )
                        {
                            area_temp = area;
                            pos_poly = h;
                        }
                        
                        //Liberer la selection
                        vec_poly[h]->close();
                    }
                    
                    //Recuperer l'identificateur de la polyline
                    poly_handle = vec_poly[pos_poly]->handle();
                    
                    //Recuperer la longueur de la polyline
                    poly_length = getLength( vec_poly[pos_poly] );
                    
                    //Recuperer la surface de la polyline
                    vec_poly[pos_poly]->getArea( poly_area );
                    
                    //Liberer la selection
                    vec_poly[pos_poly]->close();
                }
                
                //Recuperer les données
                exportText( txt, pData );
                
                // Iteration sur le tableau des données
                for( long j = 0; j < pData.size(); j++ )
                {
                    //Ecriture des données dans le fichier sélectionné
                    fichier << pData[j] << ";";
                }
                
                //Inserer le handle de la polyline
                fichier << poly_handle.ascii() << ";";
                
                //Inserer la longueur de la polyline
                fichier << to_string( poly_length ) << ";";
                
                //Inserer la surface de la polyline
                fichier << to_string( poly_area );
                
                //Se mettre à la ligne
                fichier << endl;
                
                //Incrémenter le nombre d'objet
                obj++;
                
                //Vider les données
                pData.clear();
                
                //Vider  vec_poly
                vec_poly.clear();
                vec_poly.shrink_to_fit();
                
                //Liberer la mémoire
                txt->close();
                
                //Progresser
                prog_txt.moveUp( i );
            }
            
            if( vec_error_txt.size() > 0 )
            {
                for( auto& text_error : vec_error_txt )
                {
                    //Mettre les textes dans la selection courante
                    acdbGetAdsName( ssError, text_error->id() );
                    acedSSAdd( ssError, sserror, sserror );
                    acedSSFree( ssError );
                    //Fermer le texte
                    text_error->close();
                }
            }
            
            try
            {
                //Persister les données
                fichier.close();
                
                //Recuperer la taille de vec_error_txt
                int size_error_txt = vec_error_txt.size();
                
                //Message
                print( "Exportation de " + to_string( obj ) + " textes dans " + acStrToStr( file ) + " terminée." );
                print( to_string( size_error_txt ) + " textes ignorés." );
                
                //Réinitialiser obj
                obj = 0;
                
            }
            
            catch( ... )
            {
                print( "Impossible d'enregistrer le fichier." );
                acedSSFree( ssText );
                acedSSFree( ssPoly2d );
                return;
            }
        }
        
        //Fichier excel
        else if( ext == "xlsx" )
        {
            //Initialisation du fichier excel
            workbook wb2;
            worksheet ws = wb2.active_sheet();
            
            //Ajout des en-tête du feuille
            ws.cell( 1, 1 ).value( "Handle" );
            ws.cell( 2, 1 ).value( "Calque" );
            ws.cell( 3, 1 ).value( "Couleur" );
            ws.cell( 4, 1 ).value( "Type de ligne" );
            ws.cell( 5, 1 ).value( "Echelle de type de ligne" );
            ws.cell( 6, 1 ).value( "Epaisseur de ligne" );
            ws.cell( 7, 1 ).value( "Transparence" );
            ws.cell( 8, 1 ).value( "Epaisseur" );
            ws.cell( 9, 1 ).value( "Materiau" );
            ws.cell( 10, 1 ).value( "X" );
            ws.cell( 11, 1 ).value( "Y" );
            ws.cell( 12, 1 ).value( "Z" );
            ws.cell( 13, 1 ).value( "Rotation" );
            ws.cell( 14, 1 ).value( "Valeur" );
            ws.cell( 15, 1 ).value( "Hauteur" );
            ws.cell( 16, 1 ).value( "Style" );
            ws.cell( 17, 1 ).value( "Mode verticale" );
            ws.cell( 18, 1 ).value( "Mode horizontale" );
            ws.cell( 19, 1 ).value( "Handle poly" );
            ws.cell( 20, 1 ).value( "Longueur 2d" );
            ws.cell( 21, 1 ).value( "Surface" );
            
            
            //BArre de progression
            ProgressBar prog = ProgressBar( _T( "Progression" ), size );
            int rw = 2;
            
            //Déclarer un AcDbTxt
            AcDbText* txt = NULL;
            
            for( long i = 0; i < size; i++ )
            {
                //Recuperer le i-eme texte
                txt = getTextFromSs( ssText, i, GcDb::kForRead );
                
                //Safeguard
                if( !txt )
                    continue;
                    
                //Recuperer l'extents du texte
                AcDbExtents ext;
                txt->getGeomExtents( ext );
                //Recuperer la position du texte
                AcGePoint3d al_pt = midPoint3d( ext.minPoint(), ext.maxPoint() );
                
                for( int pl = 0; pl < size_2d; pl++ )
                {
                    //Recuperer le pl-eme polyligne
                    poly_2d = getPoly2DFromSs( ssPoly2d, pl );
                    
                    //Tester si pos_2d est dans poly_2d
                    bool test_txt = isPointInPoly( poly_2d, AcGePoint2d( al_pt.x, al_pt.y ) );
                    
                    //Verifier
                    if( !test_txt )
                    {
                        poly_2d->close();
                        continue;
                    }
                    
                    //Inserer la polyline dans vec_poly
                    vec_poly.emplace_back( poly_2d );
                    
                    //Fermer la polyline
                    poly_2d->close();
                }
                
                //Déclarer un identificateur unique
                AcDbHandle poly_handle;
                
                //Longueur de la polyline
                double poly_length;
                
                //Surface de la polyline
                double poly_area;
                
                //Si aucune polyline n'entoure le text, continuer la boucle de text
                if( vec_poly.size() == 0 )
                {
                    //Inserer le texte dans la liste d'erreur
                    vec_error_txt.emplace_back( txt );
                    
                    txt->close();
                    continue;
                }
                
                //Parcourir les polylines dans vec_poly
                if( vec_poly.size() == 1 )
                {
                    //Affecter le handle de la polyline
                    poly_handle = vec_poly[0]->handle();
                    
                    //Recuperer la longueur de la polyline
                    poly_length = getLength( vec_poly[0] );
                    
                    //Recuperer la surface de la polyline
                    vec_poly[0]->getArea( poly_area );
                }
                
                else
                {
                    //Trouver la polyline qui a la plus petite surface
                    double area_temp = 1000000;
                    
                    //Position de ce polyline dans vec_poly
                    int pos_poly = 0;
                    
                    //Iterer sur vec_poly
                    for( int h = 0; h < vec_poly.size(); h++ )
                    {
                        //temp
                        double area = 0;
                        
                        //Recuperer la surface de la polyline
                        vec_poly[h]->getArea( area );
                        
                        //Recuperer la surface et la position
                        if( area < area_temp )
                        {
                            area_temp = area;
                            pos_poly = h;
                        }
                    }
                    
                    //Recuperer l'identificateur de la polyline
                    poly_handle = vec_poly[pos_poly]->handle();
                    
                    //Recuperer la longueur de la polyline
                    poly_length = getLength( vec_poly[pos_poly] );
                    
                    //Recuperer la surface de la polyline
                    vec_poly[pos_poly]->getArea( poly_area );
                }
                
                //Recuperer les données
                exportText( txt, pData );
                
                // Mettre le format des cellules en texte
                ws.columns( true ).number_format( number_format::text() );
                
                // Iteration sur les données
                for( long k = 0; k < pData.size(); k++ )
                {
                    //Changer l'encoddage des données
                    auto data = latin1_to_utf8( acStrToStr( pData[k] ) );
                    
                    //Ecriture dans le fichier excel
                    ws.cell( k + 1, rw ).value( data );
                    
                    //Inserer les informations de la polyline
                    if( k == 17 )
                    {
                        ws.cell( k + 2, rw ).value( poly_handle.ascii() );
                        ws.cell( k + 3, rw ).value( to_string( poly_length ) );
                        ws.cell( k + 4, rw ).value( to_string( poly_area ) );
                        
                    }
                }
                
                //Incrémentation du nombre d'objet
                obj++;
                
                // Vider le tableau des données
                pData.clear();
                
                //Fermer le texte
                txt->close();
                
                //Incrementer la ligne dans le fichier excel
                rw++;
                
                //Vider le vecteur des polylines
                vec_poly.clear();
                //Progresser
                prog.moveUp( i );
            }
            
            if( vec_error_txt.size() > 0 )
            {
                for( auto& text_error : vec_error_txt )
                {
                    //Mettre les textes dans la selection courante
                    acdbGetAdsName( ssError, text_error->id() );
                    acedSSAdd( ssError, sserror, sserror );
                    acedSSFree( ssError );
                    //Fermer le texte
                    text_error->close();
                }
            }
            
            try
            {
            
                //Persister les données
                wb2.save( filename );
                
                //Recuperer la taille de vec_error_txt
                int size_error_txt = vec_error_txt.size();
                
                //Affichage dans la console que les informations sont bien exportés
                print( "Exportation de " + to_string( obj ) + " textes dans " + acStrToStr( file ) + " terminée." );
                print( to_string( size_error_txt ) + " textes ignorés." );
                
                //Réinitialisation de la valeur de obj
                obj = 0;
                
            }
            
            catch( ... )
            {
                print( "Impossible d'enregistrer le fichier." );
                acedSSFree( ssText );
                acedSSFree( ssPoly2d );
                return;
            }
        }
        
        //Fichier shape
        else if( ext == "shp" )
        {
            //Lire le fichier de parametre
            vector<AcString> field = createTextPoly();
            writePrjOrNot( file );
            
            //Recuperer les points du surfaces
            vector<AcGePoint3d> point;
            map<int, vector<AcGePoint3d>> mapPoint;
            
            //Creer un fichier shp
            SHPHandle myShapeFile = SHPCreate( file, SHPT_POLYGONZ );
            
            //Creer un dbfFile
            DBFHandle dbfHandle = DBFCreate( file );
            
            //Créer un acdbobjectid
            AcDbObjectId idPoly = NULL;
            
            //Creer les champs
            int fieldNumber = createFields( dbfHandle, field );
            
            // Initialistion de la barre de progression
            ProgressBar prog = ProgressBar( _T( "Progression" ), size );
            
            double* x;    double* y;    double* z;
            int exported = 0;
            
            //Nombre de textes ignorés
            int nb_error = 0;
            
            //Boucler sur la selection
            for( int i = 0; i < size; i++ )
            {
                //Recuperer les blocks
                AcDbText* text = getTextFromSs( ssText, i, AcDb::kForRead );
                
                //surface index
                int iS; bool toDraw = false;
                int erea = INT32_MAX;
                
                //Recuperer l'extents du texte
                AcDbExtents ext;
                text->getGeomExtents( ext );
                //Recuperer la position du texte
                AcGePoint3d al_pt = midPoint3d( ext.minPoint(), ext.maxPoint() );
                
                //Boucler sur le polyligne
                for( int j = 0; j < size_2d; j++ )
                {
                    //Recuperer le polyligne
                    AcDbPolyline* poly = getPoly2DFromSs( ssPoly2d, j );
                    
                    //Verifier poly
                    if( !poly )
                        continue;
                        
                    //Verifier que le polyligne est bien fermer
                    if( poly->isClosed() )
                    {
                        //Recuperer le surface
                        double surface;
                        poly->getArea( surface );
                        surface = abs( surface );
                        
                        //Verifier si le text est a l'intérieur de la poliligne
                        if( isPointInPoly( poly, AcGePoint2d( al_pt.x, al_pt.y ) ) )
                        {
                            if( surface < erea )
                            {
                                erea = surface;
                                toDraw = true;
                                iS = j;
                                idPoly = poly->id();
                                
                                if( mapPoint.find( j ) == mapPoint.end() )
                                {
                                    //Recuperer les vertexs
                                    int vertNumber = poly->numVerts();
                                    
                                    point.resize( 0 );
                                    AcGePoint3d p;
                                    
                                    //Boucler sur les vecteurs
                                    for( int k = 0; k < vertNumber; k++ )
                                    {
                                        poly->getPointAt( k, p );
                                        point.push_back( p );
                                    }
                                    
                                    //Recuperer le point sur le sommet numero
                                    poly->getPointAt( 0, p );
                                    point.push_back( p );
                                    
                                    //Ajoute les infos dans la map
                                    mapPoint.insert( { j, point } );
                                }
                            }
                        }
                    }
                    
                    poly->close();
                }
                
                //Créer une entité
                AcDbEntity* entPoly = NULL;
                
                //Recuperer l'entité de la polyline deduit
                acdbOpenAcDbEntity( entPoly, idPoly, AcDb::kForRead );
                
                if( toDraw )
                {
                    //Recuperer le vecteur point du rang du sommets
                    vector<AcGePoint3d> vecPoint = mapPoint.at( iS );
                    drawTextPoly( text, entPoly, vecPoint, myShapeFile, dbfHandle, field );
                    exported++;
                }
                
                else
                {
                    //Mettre les textes dans la selection courante
                    acdbGetAdsName( ssError, text->id() );
                    acedSSAdd( ssError, sserror, sserror );
                    acedSSFree( ssError );
                    
                    //Incrementer le nombres des textes ignorés
                    nb_error++;
                }
                
                prog.moveUp( i );
                //Fermer le block
                text->close();
                
                // incrementer la barre de progression
                prog.moveUp( i );
                
            }
            
            //Liberer la selection
            DBFClose( dbfHandle );
            SHPClose( myShapeFile );
            
            //Affichage dans la console que les informations sont bien exportés
            print( "Exportation de " + to_string( exported ) + " textes dans " + acStrToStr( file ) + " terminée." );
            print( to_string( nb_error ) + " textes ignorés." );
            
            
        }
        
    }
    
    //Si il n'y a pas de point sélectionné
    else
        print( "Aucun texte sélectionné" );
        
        
    acedSSFree( ssPoly2d );
    acedSSFree( ssText );
    acedSSSetFirst( sserror, sserror );
    
}

void cmdExportTextPoly3D()
{
    //Déclarer les selections
    ads_name ssText;
    ads_name ssPoly3d;
    ads_name ssError, sserror;
    
    //textes ignorés
    vector<AcDbObjectId> vec_error_txt;
    
    //Fichier & Extension
    AcString file, ext;
    file = _T( "" );
    
    //Recuperer le  chemin du dwg courant
    AcString current_folder = getCurrentFileFolder();
    current_folder += _T( "\\" );
    
    //Stockage de données
    AcStringArray pData;
    
    //Poly3d qui entourent un texte
    vector<AcDb3dPolyline*> vec_poly;
    
    //Message
    print( "Veuillez séléctionner le/les texte(s) : " );
    
    //Déclaration d'un pointeur de poly_3d
    AcDb3dPolyline* poly_3d = NULL;
    
    //Nombre des textes
    long size = getSsText( ssText, "" );
    
    //Message
    print( "Veuillez séléctionner la/les polyline(s) : " );
    
    //Nombre des polylines 3d fermées
    long size_3d = getSsPoly3D( ssPoly3d, _T( "" ), true );
    
    //Vérification si il existe des textes séléctionnés
    if( size != 0 )
    {
    
        //Nombre d'objet sélectionné
        int obj = 0;
        
        //Demander ou enregistrer le fichier
        file = askForFilePath( false,
                "xlsx;shp;csv;txt",
                "Enregistrer sous",
                current_folder );
                
        //Gestion d'erreur
        if( file == _T( "" ) )
        {
            print( "Commande annulée." );
            acedSSFree( ssText );
            acedSSFree( ssPoly3d );
            return;
        }
        
        //Recuperer l'extension
        ext = getFileExt( file );
        
        //Changer l'encodage du chemin vers le fichier excel
        auto filename = latin1_to_utf8( acStrToStr( file ) );
        
        //Fichier : .txt, .csv
        if( ext == "csv" || ext == "txt" )
        {
            //Ouvrir le fichier
            fstream fichier( acStrToStr( file ), ios::app );
            
            //Ecrire les en-têtes du fichier
            fichier << "Handle;Calque;Couleur;Type de ligne;Echelle type de ligne;Epaisseur de ligne;Transparence;Epaisseur;Materiau;X;Y;Z;Rotation;Valeur;Hauteur;Style;Mode verticale;Mode horizontale;Handle poly;Longueur 2d;Longueur 3d;Surface" << endl;
            
            //Declarer un AcDbText
            AcDbText* txt;
            
            //Barre de progression
            ProgressBar prog_txt = ProgressBar( _T( "Progression:" ), size );
            
            //Iterer sur la selection de texte
            for( long i = 0; i < size; i++ )
            {
            
                //Recuperer le i-eme texte
                txt = getTextFromSs( ssText, i, GcDb::kForRead );
                
                //Safeguard
                if( !txt )
                    continue;
                    
                //Iterer sur les polylines
                for( int pl = 0; pl < size_3d; pl++ )
                {
                    //Recuperer la position du texte
                    AcGePoint3d pos_text = txt->position();
                    
                    //Recuperer le pl-eme polyline
                    poly_3d = getPoly3DFromSs( ssPoly3d, pl );
                    
                    //Tester si pos_txt est dans poly_3d
                    bool test_txt = isPointInPoly( poly_3d, pos_text );
                    
                    //Verifier
                    if( !test_txt )
                    {
                        poly_3d->close();
                        continue;
                    }
                    
                    //Inserer ce polyline dans vec_poly
                    vec_poly.emplace_back( poly_3d );
                    
                    //Liberer la mémoire
                    poly_3d->close();
                }
                
                //Identifiant unique de la polyline
                AcDbHandle poly_handle;
                
                //Longueur 2d de la polyline
                double poly_length;
                
                //Longueur 3d de la polyline
                double poly_length3d;
                
                //Surface de la polyline
                double poly_area;
                
                
                //Si aucune polyline n'entoure le text, continuer la boucle de text
                if( vec_poly.size() == 0 )
                {
                    //Recuperer les textes id
                    vec_error_txt.emplace_back( txt->id() );
                    txt->close();
                    continue;
                }
                
                //Parcourir les polylines dans vec_poly
                if( vec_poly.size() == 1 )
                {
                    poly_handle = vec_poly[0]->handle();
                    
                    //Recuperer la longueur 2d de la polyline
                    poly_length = get2DLengthPoly( vec_poly[0] );
                    
                    //Recuperer la longueur 2d de la polyline
                    poly_length3d = getLength( vec_poly[0] );
                    
                    //Recuperer la surface de la polyline
                    vec_poly[0]->getArea( poly_area );
                    
                    
                    //Liberer la selection
                    vec_poly[0]->close();
                }
                
                else
                {
                    //Surface temporaire
                    double area_temp = 1000000;
                    
                    //Position de la polyline dans vec_poly
                    int pos_poly = 0;
                    
                    //Irerer sur le vecteur vec_poly
                    for( int h = 0; h < vec_poly.size(); h++ )
                    {
                        //temp
                        double area = 0;
                        vec_poly[h]->getArea( area );
                        
                        //Recuperer la surface ainsi que la position de la polyline
                        if( area <= area_temp )
                        {
                            area_temp = area;
                            pos_poly = h;
                        }
                    }
                    
                    //Recuperer l'identifiant unique
                    poly_handle = vec_poly[pos_poly]->handle();
                    
                    //Recuperer la longueur 2d de la polyline
                    poly_length = get2DLengthPoly( vec_poly[pos_poly] );
                    
                    //Recuperer la longueur 2d de la polyline
                    poly_length3d = getLength( vec_poly[0] );
                    
                    //Recuperer la surface de la polyline
                    vec_poly[pos_poly]->getArea( poly_area );
                    
                    //Liberer la selection
                    vec_poly[pos_poly]->close();
                }
                
                //Recuperer les données
                exportText( txt, pData );
                
                //Iterer pData
                for( long j = 0; j < pData.size(); j++ )
                {
                    //Ecriture des données dans le fichier sélectionné
                    fichier << pData[j] << ";";
                }
                
                //Inserer le handle de la polyline
                fichier << poly_handle.ascii() << ";";
                
                //Inserer le handle de la polyline
                fichier << to_string( poly_length ) << ";";
                
                //Inserer le handle de la polyline
                fichier << to_string( poly_length3d ) << ";";
                
                //Inserer le handle de la polyline
                fichier << to_string( poly_area );
                
                // Se mettre à la ligne
                fichier << endl;
                
                //Incrémentation du nombre d'objet
                obj++;
                
                // Vider pData
                pData.clear();
                
                //Vider vec_poly
                vec_poly.clear();
                
                //Liberer la mémoire
                txt->close();
                
                //Progresser
                prog_txt.moveUp( i );
            }
            
            
            if( vec_error_txt.size() > 0 )
            {
                for( auto& text_error : vec_error_txt )
                {
                    //Mettre les textes dans la selection courante
                    acdbGetAdsName( ssError, text_error );
                    acedSSAdd( ssError, sserror, sserror );
                    acedSSFree( ssError );
                }
            }
            
            try
            {
                //Persister les données
                fichier.close();
                
                //Recuperer la taille de vec_error_txt
                int taille_error_txt = vec_error_txt.size();
                
                //Message
                print( "Exportation de " + to_string( obj ) + " textes dans " + acStrToStr( file ) + " terminée." );
                print( to_string( taille_error_txt ) + " textes ignorés." );
                
                //Réinitialiser obj
                obj = 0;
                
            }
            
            catch( ... )
            {
                print( "Impossible d'enregistrer le fichier." );
                acedSSFree( ssText );
                acedSSFree( ssPoly3d );
                return;
            }
        }
        
        //Fichier : .xlsx
        else if( ext == "xlsx" )
        {
            //Initialisation du fichier excel
            workbook wb2;
            worksheet ws = wb2.active_sheet();
            
            //Ajout des en-tête du feuille
            ws.cell( 1, 1 ).value( "Handle" );
            ws.cell( 2, 1 ).value( "Calque" );
            ws.cell( 3, 1 ).value( "Couleur" );
            ws.cell( 4, 1 ).value( "Type de ligne" );
            ws.cell( 5, 1 ).value( "Echelle de type de ligne" );
            ws.cell( 6, 1 ).value( "Epaisseur de ligne" );
            ws.cell( 7, 1 ).value( "Transparence" );
            ws.cell( 8, 1 ).value( "Epaisseur" );
            ws.cell( 9, 1 ).value( "Materiau" );
            ws.cell( 10, 1 ).value( "X" );
            ws.cell( 11, 1 ).value( "Y" );
            ws.cell( 12, 1 ).value( "Z" );
            ws.cell( 13, 1 ).value( "Rotation" );
            ws.cell( 14, 1 ).value( "Valeur" );
            ws.cell( 15, 1 ).value( "Hauteur" );
            ws.cell( 16, 1 ).value( "Style" );
            ws.cell( 17, 1 ).value( "Mode verticale" );
            ws.cell( 18, 1 ).value( "Mode horizontale" );
            ws.cell( 19, 1 ).value( "Handle poly" );
            ws.cell( 20, 1 ).value( "Longueur 2d" );
            ws.cell( 21, 1 ).value( "Longueur 3d" );
            ws.cell( 22, 1 ).value( "Surface" );
            
            
            //BArre de progression
            ProgressBar prog = ProgressBar( _T( "Progression" ), size );
            
            //Indice de ligne
            int rw = 2;
            
            
            //Declarer un AcDbText
            AcDbText* txt = NULL;
            
            for( long i = 0; i < size; i++ )
            {
                //Recuperer le i-eme texte
                txt = getTextFromSs( ssText, i, GcDb::kForRead );
                
                //Safeguard
                if( !txt )
                    continue;
                    
                //Iterer sur les polylines
                for( int pl = 0; pl < size_3d; pl++ )
                {
                    //Recuperer la position du texte
                    AcGePoint3d pos_text = txt->position();
                    
                    //Recuperer le pl-eme polyline
                    poly_3d = getPoly3DFromSs( ssPoly3d, pl );
                    
                    //Tester si pos_txt est dans poly_3d
                    bool test_txt = isPointInPoly( poly_3d, pos_text );
                    
                    //Verifier
                    if( !test_txt )
                    {
                        poly_3d->close();
                        continue;
                    }
                    
                    //Inserer ce polyline dans vec_poly
                    vec_poly.emplace_back( poly_3d );
                    
                    //Liberer la mémoire
                    poly_3d->close();
                }
                
                //Identifiant unique de la polyline
                AcDbHandle poly_handle;
                
                //Longueur 2d de la polyline
                double poly_length;
                
                //Longueur 2d de la polyline
                double poly_length3d;
                
                //Surface de la polyline
                double poly_area;
                
                //Si aucune polyline n'entoure le text, continuer la boucle de text
                if( vec_poly.size() == 0 )
                {
                    //Recuperer les textes ignorés
                    vec_error_txt.emplace_back( txt->id() );
                    
                    txt->close();
                    continue;
                }
                
                //Parcourir les polylines dans vec_poly
                if( vec_poly.size() == 1 )
                {
                    poly_handle = vec_poly[0]->handle();
                    
                    //Recuperer la longueur 2d de la polyline
                    poly_length = get2DLengthPoly( vec_poly[0] );
                    
                    //Recuperer la longueur 2d de la polyline
                    poly_length3d = getLength( vec_poly[0] );
                    
                    //Recuperer la surface de la polyline
                    vec_poly[0]->getArea( poly_area );
                    
                    //Liberer la mémoire
                    vec_poly[0]->close();
                }
                
                else
                {
                    //Surface temporaire
                    double area_temp = 1000000;
                    
                    //Position de la polyline dans vec_poly
                    int pos_poly = 0;
                    
                    //Irerer sur le vecteur vec_poly
                    for( int h = 0; h < vec_poly.size(); h++ )
                    {
                        //temp
                        double area = 0;
                        vec_poly[h]->getArea( area );
                        
                        //Recuperer la surface ainsi que la position de la polyline
                        if( area <= area_temp )
                        {
                            area_temp = area;
                            pos_poly = h;
                        }
                        
                        //Liberer la selection
                        vec_poly[h]->close();
                    }
                    
                    //Recuperer l'identifiant unique
                    poly_handle = vec_poly[pos_poly]->handle();
                    
                    //Recuperer la longueur 2d de la polyline
                    poly_length = get2DLengthPoly( vec_poly[pos_poly] );
                    
                    //Recuperer la longueur 2d de la polyline
                    poly_length3d = getLength( vec_poly[pos_poly] );
                    
                    //Recuperer la surface de la polyline
                    vec_poly[pos_poly]->getArea( poly_area );
                    
                    //Fermer la polyline
                    vec_poly[pos_poly]->close();
                }
                
                // Appel de la fonction de traitement
                exportText( txt, pData );
                
                // Mettre le format des cellules en texte
                ws.columns( true ).number_format( number_format::text() );
                
                // Iteration sur les données
                for( long k = 0; k < pData.size(); k++ )
                {
                    //Ecriture dans le fichier excel
                    ws.cell( k + 1, rw ).value( pData[k] );
                    
                    //Inserer le handle du polyline
                    if( k == 17 )
                    {
                        ws.cell( k + 2, rw ).value( poly_handle.ascii() );
                        ws.cell( k + 3, rw ).value( to_string( poly_length ) );
                        ws.cell( k + 4, rw ).value( to_string( poly_length3d ) );
                        ws.cell( k + 5, rw ).value( to_string( poly_area ) );
                    }
                }
                
                //Incrementer le nombre d'objet
                obj++;
                
                //Vider pData
                pData.clear();
                
                //Fermer le texte
                txt->close();
                
                //Incrementer l'indice de ligne excel
                rw++;
                
                //Vider vec_poly
                vec_poly.clear();
                
                //Progresser
                prog.moveUp( i );
            }
            
            if( vec_error_txt.size() > 0 )
            {
                for( auto& text_error : vec_error_txt )
                {
                    //Mettre les textes dans la selection courante
                    acdbGetAdsName( ssError, text_error );
                    acedSSAdd( ssError, sserror, sserror );
                    acedSSFree( ssError );
                }
            }
            
            try
            {
            
                //Persister les données
                wb2.save( filename );
                
                //Recuperer la taille de vec_error_txt
                int taille_vecerror = vec_error_txt.size();
                
                //Affichage dans la console que les informations sont bien exportés
                print( "Exportation de " + to_string( obj ) + " textes dans " + acStrToStr( file ) + " terminée." );
                print( to_string( taille_vecerror ) + " textes ignorés." );
                
                //Réinitialiser obj
                obj = 0;
            }
            
            catch( ... )
            {
                print( "Impossible d'enregistrer le fichier." );
                acedSSFree( ssText );
                acedSSFree( ssPoly3d );
                return;
            }
        }
        
        //Fichier : .shp
        else if( ext == "shp" )
        {
        
            //Lire le fichier de parametre
            vector<AcString> field = createTextPoly3d();
            writePrjOrNot( file );
            
            //Recuperer les points du surfaces
            vector<AcGePoint3d> point;
            map<int, vector<AcGePoint3d>> mapPoint;
            
            //Creer un fichier shp
            SHPHandle myShapeFile = SHPCreate( file, SHPT_POLYGONZ );
            
            //Creer un dbfFile
            DBFHandle dbfHandle = DBFCreate( file );
            
            //Creer les champs
            int fieldNumber = createFields( dbfHandle, field );
            
            // Initialistion de la barre de progression
            ProgressBar prog = ProgressBar( _T( "Progression" ), size );
            
            double* x;    double* y;    double* z;
            int exported = 0;
            
            int nb_error = 0;
            
            //Boucler sur la selection
            for( int i = 0; i < size; i++ )
            {
                //Recuperer les blocks
                AcDbText* text = getTextFromSs( ssText, i, AcDb::kForRead );
                
                //surface index
                int iS; bool toDraw = false;
                int erea = INT32_MAX;
                
                //Recuperer son point d'insertion
                AcGePoint3d ptBlock = text->position();
                
                //Créer un id
                AcDbObjectId idPoly = NULL;
                
                //Boucler sur le polyligne
                for( int j = 0; j < size_3d; j++ )
                {
                    //Recuperer le polyligne
                    AcDb3dPolyline* poly = getPoly3DFromSs( ssPoly3d, j );
                    
                    //Verifier que le polyligne est bien fermer
                    if( poly->isClosed() )
                    {
                        //Recuperer le surface
                        double surface;
                        poly->getArea( surface );
                        surface = abs( surface );
                        
                        //Verifier si le text est a l'intérieur de la polyline
                        if( isPointInPoly( poly, ptBlock ) )
                        {
                            if( surface < erea )
                            {
                                erea = surface;
                                toDraw = true;
                                iS = j;
                                
                                //Recuperer l'identifiant de la polyline
                                idPoly = poly->id();
                                
                                if( mapPoint.find( j ) == mapPoint.end() )
                                {
                                    point.resize( 0 );
                                    AcGePoint3d p;
                                    //Boucer sur les sommets
                                    AcDbObjectIterator* iter = poly->vertexIterator();
                                    AcDb3dPolylineVertex* vertex;
                                    
                                    for( iter->start(); !iter->done(); iter->step() )
                                    {
                                        if( !poly->openVertex( vertex, iter->objectId(), AcDb::kForRead ) )
                                        {
                                            p = vertex->position();
                                            point.push_back( p );
                                        }
                                    }
                                    
                                    vertex->close();
                                    //Recuperer le point sur le sommet numero 0
                                    poly->getStartPoint( p );
                                    point.push_back( p );
                                    
                                    //Ajoute les infos dans la map
                                    mapPoint.insert( { j, point } );
                                }
                            }
                        }
                    }
                    
                    //Fermer la polyline
                    poly->close();
                }
                
                //Créer une entité
                AcDbEntity* entPoly = NULL;
                
                //Recuperer l'entité de la polyline deduit
                acdbOpenAcDbEntity( entPoly, idPoly, AcDb::kForRead );
                
                if( toDraw )
                {
                    //Recuperer le vecteur point du rang du sommets
                    vector<AcGePoint3d> vecPoint = mapPoint.at( iS );
                    drawTextPoly( text, entPoly, vecPoint, myShapeFile, dbfHandle, field );
                    exported++;
                }
                
                else
                {
                    //Mettre les textes dans la selection courante
                    acdbGetAdsName( ssError, text->id() );
                    acedSSAdd( ssError, sserror, sserror );
                    acedSSFree( ssError );
                    
                    //Incrementer le nombre des textes ignorés
                    nb_error++;
                }
                
                //Progresser
                prog.moveUp( i );
                
                //Fermer le block
                text->close();
                
                // incrementer la barre de progression
                prog.moveUp( i );
            }
            
            //Liberer la memoire
            DBFClose( dbfHandle );
            SHPClose( myShapeFile );
            
            //Affichage dans la console que les informations sont bien exportés
            print( "Exportation de " + to_string( exported ) + " textes dans " + acStrToStr( file ) + " terminée." );
            print( to_string( nb_error ) + " textes ignorés." );
        }
    }
    
    //Si il n'y a pas de point sélectionné
    else
        print( "Aucun texte sélectionné" );
        
        
    acedSSFree( ssPoly3d );
    acedSSFree( ssText );
    acedSSSetFirst( sserror, sserror );
}

void cmdExportTextPoly()
{
    //Nom du texte
    ads_name ssText;
    ads_name ssPoly;
    ads_name ssError, sserror;
    
    //Fichier & Extension
    AcString file = _T( "" );
    AcString ext;
    
    //Recuperer le  chemin du dwg courant
    AcString current_folder = getCurrentFileFolder();
    current_folder += _T( "\\" );
    
    
    //Tableau de données
    AcStringArray pData;
    
    //Id des  polylines
    AcDbObjectId obj_id;
    
    //Vecteur des erreur
    vector<AcDbObjectId> vec_error;
    
    //Poly3d qui entourent un texte
    vector<AcDbEntity*> vec_entity;
    vector<double> areas;
    
    //Message
    print( "Veuillez séléctionner le/les texte(s) : " );
    
    //Déclaration d'un pointeur de poly_3d
    AcDb3dPolyline* poly_3d = NULL;
    AcDbPolyline* poly_2d = NULL;
    AcDbEntity* ent = NULL;
    
    
    //Nombre des textes
    long size = getSsText( ssText, "" );
    
    //Message
    print( "Veuillez séléctionner la/les polyline(s) : " );
    
    //Nombre des polylines 3d fermées
    long size_3d = getSelectionSet( ssPoly, _T( "" ), "LWPOLYLINE,POLYLINE" );
    
    //Vérification si il existe des textes séléctionnés
    if( size != 0 )
    {
        //Nombre d'objet sélectionné
        int obj = 0;
        
        //Demander ou enregistrer le fichier
        file = askForFilePath( false,
                "xlsx;shp;csv;txt",
                "Enregistrer sous",
                current_folder );
                
        //Gestion d'erreur
        if( file == _T( "" ) )
        {
            print( "Commande annulée." );
            acedSSFree( ssText );
            acedSSFree( ssPoly );
            return;
        }
        
        //Recuperer l'extension
        ext = getFileExt( file );
        
        //Changer l'encodage du chemin vers le fichier excel
        auto filename = latin1_to_utf8( acStrToStr( file ) );
        
        //Fichier : .txt, .csv
        if( ext == "csv" || ext == "txt" )
        {
            //Ouvrir le fichier
            fstream fichier( acStrToStr( file ), ios::app );
            
            //Ecrire les en-têtes du fichier
            fichier << "Handle;Calque;Couleur;Type de ligne;Echelle type de ligne;Epaisseur de ligne;Transparence;Epaisseur;Materiau;X;Y;Z;Rotation;Valeur;Hauteur;Style;Mode verticale;Mode horizontale;Handle poly;Longueur 2d;Longueur 3d;Surface" << endl;
            
            //Declarer un AcDbText
            AcDbText* txt;
            
            //Barre de progression
            ProgressBar prog_txt = ProgressBar( _T( "Progression:" ), size );
            
            //Iterer sur la selection de texte
            for( long i = 0; i < size; i++ )
            {
            
                //Recuperer le i-eme texte
                txt = getTextFromSs( ssText, i, GcDb::kForRead );
                
                //Safeguard
                if( !txt )
                    continue;
                    
                //Reperer le centre du texte
                AcDbExtents ext;
                txt->getGeomExtents( ext );
                
                //Recuperer son point d'insertion
                AcGePoint3d pos_text = midPoint3d( ext.minPoint(), ext.maxPoint() );
                
                //Iterer sur les polylines
                for( int pl = 0; pl < size_3d; pl++ )
                {
                    //Recuperer le pl-eme polyline
                    obj_id = getObIdFromSs( ssPoly, pl );
                    
                    //Ouvrir l'entité
                    acdbOpenAcDbEntity( ent, obj_id, AcDb::kForRead );
                    
                    //verifier ent
                    if( !ent )
                        continue;
                        
                    if( ent->isKindOf( AcDbPolyline::desc() ) )
                    {
                        //Caster l'entité en polyline
                        poly_2d = AcDbPolyline::cast( ent );
                        
                        //Verifier poly_2d
                        if( !poly_2d )
                        {
                            ent->close();
                            continue;
                        }
                        
                        if( isClosed( poly_2d ) )
                        {
                            //Tester si pos_txt est dans poly_3d
                            bool test_txt = isPointInPoly( poly_2d, AcGePoint2d( pos_text.x, pos_text.y ) );
                            
                            //Verifier
                            if( !test_txt )
                            {
                                poly_2d->close();
                                continue;
                            }
                            
                            double surf;
                            poly_2d->getArea( surf );
                            
                            //Inserer les polylines
                            vec_entity.push_back( ent );
                            areas.push_back( surf );
                            
                            //Liberer la mémoire
                            ent->close();
                        }
                    }
                    
                    else if( ent->isKindOf( AcDb3dPolyline::desc() ) )
                    {
                        //Caster l'entité en polyline
                        poly_3d = AcDb3dPolyline::cast( ent );
                        
                        //Verifier poly_2d
                        if( !poly_3d )
                        {
                            ent->close();
                            continue;
                        }
                        
                        if( isClosed( poly_3d ) )
                        {
                        
                            //Tester si pos_txt est dans poly_3d
                            bool test_txt = isPointInPoly( poly_3d, pos_text );
                            
                            //Verifier
                            if( !test_txt )
                            {
                                poly_3d->close();
                                continue;
                            }
                            
                            double surf;
                            poly_3d->getArea( surf );
                            
                            //Inserer les polylines
                            vec_entity.push_back( ent );
                            areas.push_back( surf );
                            
                            //Liberer la mémoire
                            ent->close();
                        }
                        
                        
                    }
                    
                    //Liberer la selection
                    ent->close();
                    
                }
                
                //Identifiant unique de la polyline
                AcDbHandle poly_handle;
                AcDbPolyline* _temp_poly = NULL;
                AcDb3dPolyline* _temp_poly3d = NULL;
                
                //Longueur 2d de la polyline
                double poly_length;
                
                //Longueur 3d de la polyline
                double poly_length3d;
                
                //Surface de la polyline
                double poly_area;
                
                
                //Si aucune polyline n'entoure le text, continuer la boucle de text
                if( vec_entity.size() == 0 )
                {
                    txt->close();
                    
                    vec_error.emplace_back( txt->id() );
                    
                    //Progresser
                    prog_txt.moveUp( i );
                    continue;
                }
                
                //Parcourir les polylines dans vec_poly
                if( vec_entity.size() == 1 )
                {
                
                    //verifier _ent
                    if( !vec_entity[0] )
                    {
                        vec_error.emplace_back( txt->id() );
                        
                        txt->close();
                        prog_txt.moveUp( i );
                        continue;
                    }
                    
                    //Affecter le handle de la polyline
                    poly_handle = vec_entity[0]->handle();
                    
                    
                    //Caster l'entitié en polyline
                    if( ent->isKindOf( AcDbPolyline::desc() ) )
                    {
                        AcDbPolyline* ptemp = AcDbPolyline::cast( vec_entity[0] );
                        
                        //verifier ptemp
                        if( !ptemp )
                        {
                            //Liberer la mémoire
                            vec_entity[0]->close();
                            txt->close();
                            prog_txt.moveUp( i );
                            continue;
                        }
                        
                        //Recuperer la longueur de la polyline
                        poly_length = getLength( ptemp );
                        
                        poly_length3d = poly_length;
                        
                        //Recuperer la surface de la polyline
                        ptemp->getArea( poly_area );
                        
                        //Liberer la mémoire
                        vec_entity[0]->close();
                    }
                    
                    else if( ent->isKindOf( AcDb3dPolyline::desc() ) )
                    {
                        AcDb3dPolyline* ptemp = AcDb3dPolyline::cast( vec_entity[0] );
                        
                        //verifier ptemp
                        if( !ptemp )
                        {
                            //Liberer la mémoire
                            vec_entity[0]->close();
                            txt->close();
                            prog_txt.moveUp( i );
                            continue;
                        }
                        
                        //Recuperer la longueur de la polyline
                        poly_length = getLength( ptemp );
                        
                        //Recuperer la longueur 3d de la polyline
                        poly_length3d = get2DLengthPoly( ptemp );
                        
                        //Recuperer la surface de la polyline
                        ptemp->getArea( poly_area );
                        
                        vec_entity[0]->close();
                    }
                    
                }
                
                else
                {
                    //Trouver la polyline qui a la plus petite surface
                    double area_temp = 10000000000;
                    
                    //Position de ce polyline dans vec_poly
                    int pos_poly = 0;
                    
                    //Iterer sur vec_poly
                    for( int h = 0; h < vec_entity.size(); h++ )
                    {
                        //temp
                        double area = 0;
                        
                        area = areas[h];
                        
                        //Recuperer la surface et la position
                        if( area < area_temp )
                        {
                            area_temp = area;
                            pos_poly = h;
                        }
                    }
                    
                    //Recuperer l'identificateur de la polyline
                    poly_handle = vec_entity[pos_poly]->handle();
                    
                    if( vec_entity[pos_poly]->isKindOf( AcDbPolyline::desc() ) )
                    {
                        _temp_poly = AcDbPolyline::cast( vec_entity[pos_poly] );
                        
                        //verifier _temp_poly
                        if( !_temp_poly )
                            continue;
                            
                        //Recuperer la longueur de la polyline
                        poly_length = getLength( _temp_poly );
                        
                        poly_length3d = poly_length;
                        
                        //Recuperer la surface de la polyline
                        _temp_poly->getArea( poly_area );
                        
                        vec_entity[pos_poly]->close();
                    }
                    
                    else if( vec_entity[pos_poly]->isKindOf( AcDb3dPolyline::desc() ) )
                    {
                        _temp_poly3d = AcDb3dPolyline::cast( vec_entity[pos_poly] );
                        
                        
                        //verifier _temp_poly3d
                        if( !_temp_poly3d )
                            continue;
                            
                        //Recuperer la longueur de la polyline
                        poly_length = getLength( _temp_poly3d );
                        
                        //Recuperer la longueur 3d de la polyline
                        poly_length3d = get2DLengthPoly( _temp_poly3d );
                        
                        //Recuperer la surface de la polyline
                        _temp_poly3d->getArea( poly_area );
                        
                        //Liberer la mémoire
                        vec_entity[pos_poly]->close();
                    }
                    
                }
                
                //Recuperer les données
                exportText( txt, pData );
                
                //Iterer pData
                for( long j = 0; j < pData.size(); j++ )
                {
                    //Ecriture des données dans le fichier sélectionné
                    fichier << pData[j] << ";";
                }
                
                //Inserer le handle de la polyline
                fichier << poly_handle.ascii() << ";";
                
                //Inserer le handle de la polyline
                fichier << to_string( poly_length ) << ";";
                
                //Inserer le handle de la polyline
                fichier << to_string( poly_length3d ) << ";";
                
                //Inserer le handle de la polyline
                fichier << to_string( poly_area );
                
                // Se mettre à la ligne
                fichier << endl;
                
                //Incrémentation du nombre d'objet
                obj++;
                
                // Vider pData
                pData.clear();
                
                //Vider vec_poly
                vec_entity.clear();
                
                vec_entity.shrink_to_fit();
                
                //Liberer la mémoire
                txt->close();
                
                //Progresser
                prog_txt.moveUp( i );
            }
            
            //Recuperer le nombre de blocks ignorés
            int taille_txt_ignored = vec_error.size();
            
            if( taille_txt_ignored > 0 )
            {
                for( auto& obj_id : vec_error )
                {
                    acdbGetAdsName( ssError, obj_id );
                    acedSSAdd( ssError, sserror, sserror );
                    acedSSFree( ssError );
                }
            }
            
            try
            {
                //Persister les données
                fichier.close();
                
                //Message
                print( "Exportation de " + to_string( obj ) + " textes dans " + acStrToStr( file ) + " terminée." );
                print( "Nombre de textes ignorés: " + to_string( taille_txt_ignored ) );
                //Réinitialiser obj
                obj = 0;
                
            }
            
            catch( ... )
            {
                print( "Impossible d'enregistrer le fichier." );
                acedSSFree( ssText );
                acedSSFree( ssPoly );
                return;
            }
        }
        
        //Fichier : .xlsx
        else if( ext == "xlsx" )
        {
            //Initialisation du fichier excel
            workbook wb2;
            worksheet ws = wb2.active_sheet();
            
            //Ajout des en-tête du feuille
            ws.cell( 1, 1 ).value( "Handle" );
            ws.cell( 2, 1 ).value( "Calque" );
            ws.cell( 3, 1 ).value( "Couleur" );
            ws.cell( 4, 1 ).value( "Type de ligne" );
            ws.cell( 5, 1 ).value( "Echelle de type de ligne" );
            ws.cell( 6, 1 ).value( "Epaisseur de ligne" );
            ws.cell( 7, 1 ).value( "Transparence" );
            ws.cell( 8, 1 ).value( "Epaisseur" );
            ws.cell( 9, 1 ).value( "Materiau" );
            ws.cell( 10, 1 ).value( "X" );
            ws.cell( 11, 1 ).value( "Y" );
            ws.cell( 12, 1 ).value( "Z" );
            ws.cell( 13, 1 ).value( "Rotation" );
            ws.cell( 14, 1 ).value( "Valeur" );
            ws.cell( 15, 1 ).value( "Hauteur" );
            ws.cell( 16, 1 ).value( "Style" );
            ws.cell( 17, 1 ).value( "Mode verticale" );
            ws.cell( 18, 1 ).value( "Mode horizontale" );
            ws.cell( 19, 1 ).value( "Handle poly" );
            ws.cell( 20, 1 ).value( "Longueur 2d" );
            ws.cell( 21, 1 ).value( "Longueur 3d" );
            ws.cell( 22, 1 ).value( "Surface" );
            
            
            //BArre de progression
            ProgressBar prog = ProgressBar( _T( "Progression" ), size );
            
            //Indice de ligne
            int rw = 2;
            
            //Declarer un AcDbText
            AcDbText* txt = NULL;
            
            for( long i = 0; i < size; i++ )
            {
                //Recuperer le i-eme texte
                txt = getTextFromSs( ssText, i, GcDb::kForRead );
                
                //Safeguard
                if( !txt )
                    continue;
                    
                //Reperer le centre du texte
                AcDbExtents ext;
                txt->getGeomExtents( ext );
                
                //Recuperer son point d'insertion
                AcGePoint3d pos_text = midPoint3d( ext.minPoint(), ext.maxPoint() );
                
                //Iterer sur les polylines
                for( int pl = 0; pl < size_3d; pl++ )
                {
                    //Recuperer le pl-eme polyline
                    obj_id = getObIdFromSs( ssPoly, pl );
                    
                    //Ouvrir l'entité
                    acdbOpenAcDbEntity( ent, obj_id, AcDb::kForRead );
                    
                    //verifier ent
                    if( !ent )
                        continue;
                        
                    if( ent->isKindOf( AcDbPolyline::desc() ) )
                    {
                        //Caster l'entité en polyline
                        poly_2d = AcDbPolyline::cast( ent );
                        
                        //Verifier poly_2d
                        if( !poly_2d )
                        {
                            ent->close();
                            continue;
                        }
                        
                        if( isClosed( poly_2d ) )
                        {
                            //Tester si pos_txt est dans poly_3d
                            bool test_txt = isPointInPoly( poly_2d, AcGePoint2d( pos_text.x, pos_text.y ) );
                            
                            //Verifier
                            if( !test_txt )
                            {
                                poly_2d->close();
                                continue;
                            }
                            
                            double surf;
                            poly_2d->getArea( surf );
                            
                            //Inserer les polylines
                            vec_entity.push_back( ent );
                            areas.push_back( surf );
                            
                            //Liberer la mémoire
                            ent->close();
                        }
                    }
                    
                    else if( ent->isKindOf( AcDb3dPolyline::desc() ) )
                    {
                        //Caster l'entité en polyline
                        poly_3d = AcDb3dPolyline::cast( ent );
                        
                        //Verifier poly_2d
                        if( !poly_3d )
                        {
                            ent->close();
                            continue;
                        }
                        
                        if( isClosed( poly_3d ) )
                        {
                        
                            //Tester si pos_txt est dans poly_3d
                            bool test_txt = isPointInPoly( poly_3d, pos_text );
                            
                            //Verifier
                            if( !test_txt )
                            {
                                poly_3d->close();
                                continue;
                            }
                            
                            double surf;
                            poly_3d->getArea( surf );
                            
                            //Inserer les polylines
                            vec_entity.push_back( ent );
                            areas.push_back( surf );
                            
                            //Liberer la mémoire
                            ent->close();
                        }
                        
                        
                    }
                    
                    //Liberer la selection
                    ent->close();
                    
                }
                
                //Identifiant unique de la polyline
                AcDbHandle poly_handle;
                AcDbPolyline* _temp_poly = NULL;
                AcDb3dPolyline* _temp_poly3d = NULL;
                
                //Longueur 2d de la polyline
                double poly_length;
                
                //Longueur 3d de la polyline
                double poly_length3d;
                
                //Surface de la polyline
                double poly_area;
                
                
                //Si aucune polyline n'entoure le text, continuer la boucle de text
                if( vec_entity.size() == 0 )
                {
                    txt->close();
                    
                    vec_error.emplace_back( txt->id() );
                    
                    //Progresser
                    prog.moveUp( i );
                    continue;
                }
                
                //Parcourir les polylines dans vec_poly
                if( vec_entity.size() == 1 )
                {
                
                    //verifier _ent
                    if( !vec_entity[0] )
                    {
                        vec_error.emplace_back( txt->id() );
                        
                        txt->close();
                        prog.moveUp( i );
                        continue;
                    }
                    
                    //Affecter le handle de la polyline
                    poly_handle = vec_entity[0]->handle();
                    
                    
                    //Caster l'entitié en polyline
                    if( ent->isKindOf( AcDbPolyline::desc() ) )
                    {
                        AcDbPolyline* ptemp = AcDbPolyline::cast( vec_entity[0] );
                        
                        //verifier ptemp
                        if( !ptemp )
                        {
                            //Liberer la mémoire
                            vec_entity[0]->close();
                            txt->close();
                            prog.moveUp( i );
                            continue;
                        }
                        
                        //Recuperer la longueur de la polyline
                        poly_length = getLength( ptemp );
                        
                        poly_length3d = poly_length;
                        
                        //Recuperer la surface de la polyline
                        ptemp->getArea( poly_area );
                        
                        //Liberer la mémoire
                        vec_entity[0]->close();
                    }
                    
                    else if( ent->isKindOf( AcDb3dPolyline::desc() ) )
                    {
                        AcDb3dPolyline* ptemp = AcDb3dPolyline::cast( vec_entity[0] );
                        
                        //verifier ptemp
                        if( !ptemp )
                        {
                            //Liberer la mémoire
                            vec_entity[0]->close();
                            txt->close();
                            prog.moveUp( i );
                            continue;
                        }
                        
                        //Recuperer la longueur de la polyline
                        poly_length = getLength( ptemp );
                        
                        //Recuperer la longueur 3d de la polyline
                        poly_length3d = get2DLengthPoly( ptemp );
                        
                        //Recuperer la surface de la polyline
                        ptemp->getArea( poly_area );
                        
                        vec_entity[0]->close();
                    }
                    
                }
                
                else
                {
                    //Trouver la polyline qui a la plus petite surface
                    double area_temp = 10000000000;
                    
                    //Position de ce polyline dans vec_poly
                    int pos_poly = 0;
                    
                    //Iterer sur vec_poly
                    for( int h = 0; h < vec_entity.size(); h++ )
                    {
                        //temp
                        double area = 0;
                        
                        area = areas[h];
                        
                        //Recuperer la surface et la position
                        if( area < area_temp )
                        {
                            area_temp = area;
                            pos_poly = h;
                        }
                    }
                    
                    //Recuperer l'identificateur de la polyline
                    poly_handle = vec_entity[pos_poly]->handle();
                    
                    if( vec_entity[pos_poly]->isKindOf( AcDbPolyline::desc() ) )
                    {
                        _temp_poly = AcDbPolyline::cast( vec_entity[pos_poly] );
                        
                        //verifier _temp_poly
                        if( !_temp_poly )
                            continue;
                            
                        //Recuperer la longueur de la polyline
                        poly_length = getLength( _temp_poly );
                        
                        poly_length3d = poly_length;
                        
                        //Recuperer la surface de la polyline
                        _temp_poly->getArea( poly_area );
                        
                        vec_entity[pos_poly]->close();
                    }
                    
                    else if( vec_entity[pos_poly]->isKindOf( AcDb3dPolyline::desc() ) )
                    {
                        _temp_poly3d = AcDb3dPolyline::cast( vec_entity[pos_poly] );
                        
                        
                        //verifier _temp_poly3d
                        if( !_temp_poly3d )
                            continue;
                            
                        //Recuperer la longueur de la polyline
                        poly_length = getLength( _temp_poly3d );
                        
                        //Recuperer la longueur 3d de la polyline
                        poly_length3d = get2DLengthPoly( _temp_poly3d );
                        
                        //Recuperer la surface de la polyline
                        _temp_poly3d->getArea( poly_area );
                        
                        //Liberer la mémoire
                        vec_entity[pos_poly]->close();
                    }
                    
                }
                
                // Appel de la fonction de traitement
                exportText( txt, pData );
                
                // Mettre le format des cellules en texte
                ws.columns( true ).number_format( number_format::text() );
                
                // Iteration sur les données
                for( long k = 0; k < pData.size(); k++ )
                {
                    //Ecriture dans le fichier excel
                    ws.cell( k + 1, rw ).value( pData[k] );
                    
                    //Inserer le handle du polyline
                    if( k == 17 )
                    {
                        ws.cell( k + 2, rw ).value( poly_handle.ascii() );
                        ws.cell( k + 3, rw ).value( to_string( poly_length ) );
                        ws.cell( k + 4, rw ).value( to_string( poly_length3d ) );
                        ws.cell( k + 5, rw ).value( to_string( poly_area ) );
                    }
                }
                
                //Incrementer le nombre d'objet
                obj++;
                
                //Vider pData
                pData.clear();
                
                //Fermer le texte
                txt->close();
                
                //Incrementer l'indice de ligne excel
                rw++;
                
                //Vider vec_poly
                vec_entity.clear();
                
                //Progresser
                prog.moveUp( i );
            }
            
            //Recuperer le nombre de blocks ignorés
            int taille_txt_ignored = vec_error.size();
            
            if( taille_txt_ignored > 0 )
            {
                for( auto& obj_id : vec_error )
                {
                    acdbGetAdsName( ssError, obj_id );
                    acedSSAdd( ssError, sserror, sserror );
                    acedSSFree( ssError );
                }
            }
            
            try
            {
            
                //Persister les données
                wb2.save( filename );
                
                //Affichage dans la console que les informations sont bien exportés
                print( "Exportation de " + to_string( obj ) + " textes dans " + acStrToStr( file ) + " terminée." );
                print( "Nombre de textes ignorés: " + to_string( taille_txt_ignored ) );
                //Réinitialiser obj
                obj = 0;
            }
            
            catch( ... )
            {
                print( "Impossible d'enregistrer le fichier." );
                acedSSFree( ssText );
                acedSSFree( ssPoly );
                return;
            }
        }
        
        //Fichier : .shp
        else if( ext == "shp" )
        {
        
            //Lire le fichier de parametre
            vector<AcString> field = createTextPoly3d();
            writePrjOrNot( file );
            
            //Recuperer les points du surfaces
            vector<AcGePoint3d> point;
            map<int, vector<AcGePoint3d>> mapPoint;
            
            //Creer un fichier shp
            SHPHandle myShapeFile = SHPCreate( file, SHPT_POLYGONZ );
            
            //Creer un dbfFile
            DBFHandle dbfHandle = DBFCreate( file );
            
            //Creer les champs
            int fieldNumber = createFields( dbfHandle, field );
            
            // Initialistion de la barre de progression
            ProgressBar prog = ProgressBar( _T( "Progression" ), size );
            
            double* x;    double* y;    double* z;
            int exported = 0;
            
            int nb_error = 0;
            
            //Boucler sur la selection
            for( int i = 0; i < size; i++ )
            {
                //Recuperer les blocks
                AcDbText* text = getTextFromSs( ssText, i, AcDb::kForRead );
                
                //surface index
                int iS; bool toDraw = false;
                int erea = INT32_MAX;
                
                //Reperer le centre du texte
                AcDbExtents ext;
                text->getGeomExtents( ext );
                
                //Recuperer son point d'insertion
                AcGePoint3d pPos = midPoint3d( ext.minPoint(), ext.maxPoint() );
                
                //Créer un id
                AcDbObjectId idPoly = NULL;
                
                //Boucler sur le polyligne
                for( int j = 0; j < size_3d; j++ )
                {
                    //Recuperer le j-eme polyline
                    obj_id = getObIdFromSs( ssPoly, j );
                    
                    //Ouvrir l'entité
                    acdbOpenAcDbEntity( ent, obj_id, AcDb::kForRead );
                    
                    //verifier ent
                    if( !ent )
                        continue;
                        
                    if( ent->isKindOf( AcDb3dPolyline::desc() ) )
                    {
                        //Recuperer le polyligne
                        AcDb3dPolyline* poly = AcDb3dPolyline::cast( ent );
                        
                        //verifier poly
                        if( !poly )
                        {
                            ent->close();
                            continue;
                        }
                        
                        //Verifier que le polyligne est bien fermer
                        if( poly->isClosed() )
                        {
                            //Recuperer le surface
                            double surface;
                            poly->getArea( surface );
                            surface = abs( surface );
                            
                            //Verifier si le block est a l'itnerirur du poliligne
                            if( isPointInPoly( poly, pPos ) )
                            {
                                if( surface < erea )
                                {
                                    erea = surface;
                                    toDraw = true;
                                    iS = j;
                                    
                                    //Recuperer l'identifiant de la polyline
                                    idPoly = poly->id();
                                    
                                    if( mapPoint.find( j ) == mapPoint.end() )
                                    {
                                        point.resize( 0 );
                                        AcGePoint3d p;
                                        //Boucer sur les sommets
                                        AcDbObjectIterator* iter = poly->vertexIterator();
                                        AcDb3dPolylineVertex* vertex;
                                        
                                        for( iter->start(); !iter->done(); iter->step() )
                                        {
                                            if( !poly->openVertex( vertex, iter->objectId(), AcDb::kForRead ) )
                                            {
                                                p = vertex->position();
                                                point.push_back( p );
                                            }
                                        }
                                        
                                        vertex->close();
                                        //Recuperer le point sur le sommet numero 0
                                        poly->getStartPoint( p );
                                        point.push_back( p );
                                        
                                        //Ajoute les infos dans la map
                                        mapPoint.insert( { j, point } );
                                    }
                                }
                            }
                        }
                        
                        //Fermer la polyline
                        ent->close();
                    }
                    
                    else if( ent->isKindOf( AcDbPolyline::desc() ) )
                    {
                        //Recuperer le polyligne
                        AcDbPolyline* poly = AcDbPolyline::cast( ent );
                        
                        //Verifier poly
                        if( !poly )
                        {
                            ent->close();
                            continue;
                        }
                        
                        //Verifier que le polyligne est bien fermer
                        if( poly->isClosed() )
                        {
                            //Recuperer le surface
                            double surface;
                            poly->getArea( surface );
                            surface = abs( surface );
                            
                            //Verifier si le text est a l'intérieur de la poliligne
                            if( isPointInPoly( poly, AcGePoint2d( pPos.x, pPos.y ) ) )
                            {
                                if( surface < erea )
                                {
                                    erea = surface;
                                    toDraw = true;
                                    iS = j;
                                    idPoly = poly->id();
                                    
                                    if( mapPoint.find( j ) == mapPoint.end() )
                                    {
                                        //Recuperer les vertexs
                                        int vertNumber = poly->numVerts();
                                        
                                        point.resize( 0 );
                                        AcGePoint3d p;
                                        
                                        //Boucler sur les vecteurs
                                        for( int k = 0; k < vertNumber; k++ )
                                        {
                                            poly->getPointAt( k, p );
                                            point.push_back( p );
                                        }
                                        
                                        //Recuperer le point sur le sommet numero
                                        poly->getPointAt( 0, p );
                                        point.push_back( p );
                                        
                                        //Ajoute les infos dans la map
                                        mapPoint.insert( { j, point } );
                                    }
                                }
                            }
                        }
                        
                        ent->close();
                    }
                    
                }
                
                //Créer une entité
                AcDbEntity* entPoly = NULL;
                
                //Recuperer l'entité de la polyline deduit
                acdbOpenAcDbEntity( entPoly, idPoly, AcDb::kForRead );
                
                if( toDraw && entPoly )
                {
                    //Recuperer le vecteur point du rang du sommets
                    vector<AcGePoint3d> vecPoint = mapPoint.at( iS );
                    drawTextPoly( text, entPoly, vecPoint, myShapeFile, dbfHandle, field );
                    exported++;
                    
                    entPoly->close();
                }
                
                else
                {
                    acdbGetAdsName( ssError, text->id() );
                    acedSSAdd( ssError, sserror, sserror );
                    acedSSFree( ssError );
                    
                    //Incrementer les nombre d'erreurs
                    nb_error++;
                }
                
                //Progresser
                prog.moveUp( i );
                
                //Fermer le block
                text->close();
                
                // incrementer la barre de progression
                prog.moveUp( i );
            }
            
            //Liberer la memoire
            DBFClose( dbfHandle );
            SHPClose( myShapeFile );
            
            //Affichage dans la console que les informations sont bien exportés
            print( "Exportation de " + to_string( exported ) + " textes dans " + acStrToStr( file ) + " terminée." );
            print( "Nombre de textes ignorés: " + to_string( nb_error ) );
            
            
        }
        
    }
    
    //Si il n'y a pas de point sélectionné
    else
        print( "Aucun texte sélectionné" );
        
        
    //Liberer la mémoire
    acedSSFree( ssText );
    acedSSFree( ssPoly );
    acedSSSetFirst( sserror, sserror );
    
}

void cmdExportPonctuel()
{

    //Définir la selection
    ads_name ssPonct;
    
    //Fichier & extension
    AcString file = _T( "" );
    AcString ext;
    
    //Recuperer le  chemin du dwg courant
    AcString current_folder = getCurrentFileFolder();
    current_folder += _T( "\\" );
    
    //Tableau de stockage de données
    AcStringArray pData;
    
    
    //Demander au dessinateur de faire une selection
    print( "Veuillez séléctionner le/les point(s) : " );
    
    //Nombre de points dans la selection
    long size = getSsObject( ssPonct, _T( "" ) );
    
    //long size = getSelectionSet( ssPonct, "", _T( "POINT,CIRCLE,BLOCK REFERENCE,TEXT,MTEXT" ) );
    
    //Verifier le nombre de points
    if( size != 0 )
    {
    
        //Nombre d'objet sélectionné
        int obj = 0;
        
        //Initialiser file
        file = _T( "" );
        
        //Ou enregistrer le fichier?
        file = askForFilePath( false,
                "xlsx;shp;csv;txt",
                "Enregistrer sous",
                current_folder );
                
        //Gestion d'erreur
        if( file == _T( "" ) )
        {
            print( "Commande annulée." );
            acedSSFree( ssPonct );
            return;
        }
        
        //Recuperer l'extension du fichier
        ext = getFileExt( file );
        
        
        //Changer l'encodage du chemin vers le fichier excel en utf8
        auto filename = latin1_to_utf8( acStrToStr( file ) );
        
        //Si : .txt, .csv
        if( ext == "csv" || ext == "txt" )
        {
            //Ouvrir le fichier
            fstream fichier( acStrToStr( file ), ios::app );
            
            //Ecrire les en-têtes du fichier
            fichier << "Handle;Calque;Couleur;Type de ligne;Echelle de type de ligne;Epaisseur de ligne;Transparence;Epaisseur;Materiau;X;Y;Z" << endl;
            
            //Déclarer un AcDbPoint
            AcDbObjectId obj_id;
            AcDbEntity* ent = NULL;
            
            //Barre de progression
            ProgressBar prog_txt = ProgressBar( _T( "Progression:" ), size );
            
            //Iterer sur la selection de points
            for( long i = 0; i < size; i++ )
            {
            
                //Le pointeur sur le point
                obj_id = getObIdFromSs( ssPonct, i );
                
                //Ouvrir l'entité
                acdbOpenAcDbEntity( ent, obj_id, AcDb::kForRead );
                
                //verifier ent
                if( !ent )
                    continue;
                    
                    
                if(
                    ent->isKindOf( AcDbText::desc() )
                    || ent->isKindOf( AcDbMText::desc() )
                    || ent->isKindOf( AcDbBlockReference::desc() )
                    || ent->isKindOf( AcDbCircle::desc() )
                    || ent->isKindOf( AcDbPoint::desc() ) )
                    
                {
                    // Appel de la fonction de traitement
                    exportEntity( ent, pData );
                }
                
                else
                {
                    //Fermer l'entité
                    ent->close();
                    continue;
                }
                
                // Iteration sur le tableau des données
                for( long j = 0; j < pData.size(); j++ )
                {
                    //Ecriture des données dans le fichier sélectionné
                    fichier << pData[j] << ";";
                }
                
                // Se mettre à la ligne
                fichier << endl;
                
                //Incrémentation du nombre d'objet
                obj++;
                
                //Reset pData
                pData.clear();
                
                //Liberer la mémoire
                ent->close();
                
                //Progresser
                prog_txt.moveUp( i );
            }
            
            try
            {
                //Fermeture du fichier
                fichier.close();
                
                //Affichage dans la console que les informations sont bien exportés
                print( "Exportation de " + to_string( obj ) + " ponctuel(s) dans " + acStrToStr( file ) + " terminée." );
                
            }
            
            catch( ... )
            {
                print( "Impossible d'enregistrer le fichier." );
                acedSSFree( ssPonct );
                return;
            }
        }
        
        //Fichier excel
        else if( ext == "xlsx" )
        {
            //Initialisation du fichier excel
            workbook wb2;
            worksheet ws = wb2.active_sheet();
            
            //Ajout des en-tête du feuille
            ws.cell( 1, 1 ).value( "Handle" );
            ws.cell( 2, 1 ).value( "Calque" );
            ws.cell( 3, 1 ).value( "Couleur" );
            ws.cell( 4, 1 ).value( "Type de ligne" );
            ws.cell( 5, 1 ).value( "Echelle de type de ligne" );
            ws.cell( 6, 1 ).value( "Epaisseur de ligne" );
            ws.cell( 7, 1 ).value( "Transparence" );
            ws.cell( 8, 1 ).value( "Epaisseur" );
            ws.cell( 9, 1 ).value( "Materiau" );
            ws.cell( 10, 1 ).value( "X" );
            ws.cell( 11, 1 ).value( "Y" );
            ws.cell( 12, 1 ).value( "Z" );
            
            //Déclarer un AcDbPoint
            AcDbObjectId obj_id;
            AcDbEntity* ent = NULL;
            
            //BArre de progression
            ProgressBar prog = ProgressBar( _T( "Progression" ), size );
            
            int cell_cp = 0;
            
            for( long i = 0; i < size; i++ )
            {
            
                //Recuperer le i-eme objectid
                obj_id = getObIdFromSs( ssPonct, i );
                
                //Ouvrir l'entité
                acdbOpenAcDbEntity( ent, obj_id, AcDb::kForRead );
                
                //verifier ent
                if( !ent )
                    continue;
                    
                if(
                    ent->isKindOf( AcDbText::desc() )
                    || ent->isKindOf( AcDbMText::desc() )
                    || ent->isKindOf( AcDbBlockReference::desc() )
                    || ent->isKindOf( AcDbCircle::desc() )
                    || ent->isKindOf( AcDbPoint::desc() ) )
                    
                {
                    // Appel de la fonction de traitement
                    exportEntity( ent, pData );
                    
                }
                
                else
                {
                    //Fermer l'entité
                    ent->close();
                    continue;
                }
                
                // Mettre le format des cellules en texte
                ws.columns( true ).number_format( number_format::text() );
                
                // Iteration sur les données
                for( long k = 0; k < pData.size(); k++ )
                {
                    //Ecriture dans le fichier excel
                    ws.cell( k + 1, cell_cp + 2 ).value( pData[k] );
                }
                
                //Incrémentation du nombre d'objet
                obj++;
                cell_cp++;
                
                // Vider le tableau des données
                pData.clear();
                
                //Fermer le point
                ent->close();
                
                //Progresser
                prog.moveUp( i );
            }
            
            try
            {
                //Enregistrer les modifications
                wb2.save( filename );
                
                //Affichage dans la console que les informations sont bien exportés
                print( "Exportation de " + to_string( obj ) + " ponctuel(s) dans " + acStrToStr( file ) + " terminée." );
                
                //Réinitialisation de la valeur de obj
                obj = 0;
            }
            
            catch( ... )
            {
                print( "Impossible d'enregistrer le fichier." );
                acedSSFree( ssPonct );
                return;
            }
        }
        
        //Fichier shape
        else if( ext == "shp" )
        {
            //Lire le fichier de parametre
            vector<AcString> field = createPointField();
            writePrjOrNot( file );
            
            //Creer un fichier shp
            SHPHandle myShapeFile = SHPCreate( file, SHPT_POINTZ );
            
            //Déclarer un AcDbPoint
            AcDbObjectId obj_id;
            AcDbEntity* ent = NULL;
            
            //Creer un dbfFile
            DBFHandle dbfHandle = DBFCreate( file );
            
            //Creer les champs
            int fieldNumber = createFields( dbfHandle, field );
            
            // Initialistion de la barre de progression
            ProgressBar prog = ProgressBar( _T( "Progression" ), size );
            
            //Nombre de points exportés
            int nb_el = 0;
            
            //Boucler sur toute les blocs
            for( int i = 0; i < size; i++ )
            {
                //Recuperer le i-eme objectid
                obj_id = getObIdFromSs( ssPonct, i );
                
                //Ouvrir l'entité
                acdbOpenAcDbEntity( ent, obj_id, AcDb::kForRead );
                
                //verifier ent
                if( !ent )
                    continue;
                    
                    
                if(
                    ent->isKindOf( AcDbText::desc() )
                    || ent->isKindOf( AcDbMText::desc() )
                    || ent->isKindOf( AcDbBlockReference::desc() )
                    || ent->isKindOf( AcDbCircle::desc() )
                    || ent->isKindOf( AcDbPoint::desc() ) )
                    
                {
                    //Exporter les points
                    drawPonctuel( ent, myShapeFile, dbfHandle, field );
                }
                
                else
                {
                    //Fermer l'entité
                    ent->close();
                    continue;
                }
                
                
                
                
                //Incrementer le nombre de points éxportés
                nb_el++;
                
                //Liberer la mémoire
                ent->close();
                
                //Progresser
                prog.moveUp( i );
            }
            
            
            //Liberer la mémoire
            DBFClose( dbfHandle );
            SHPClose( myShapeFile );
            
            // Messages
            print( "Exportation de " + to_string( nb_el ) + " points dans " + acStrToStr( file ) + " terminée." );
        }
    }
    
    //Si il n'y a pas de point sélectionné
    else
        print( "Aucun point sélectionné" );
        
        
    //Liberer la mémoire
    acedSSFree( ssPonct );
    
}

void cmdExportHatch()
{
    //Définir la selection
    ads_name ssHatch;
    
    // Le fichier et l'extension
    AcString file,
             ext;
             
    file = _T( "" );
    
    //Recuperer le  chemin du dwg courant
    AcString current_folder = getCurrentFileFolder();
    current_folder += _T( "\\" );
    
    // Tableau de stockage des données
    AcStringArray pData;
    
    //Demander au dessinateur de selectionner
    print( "Veuillez séléctionner le/les hachure(s) : " );
    
    //Nombre des objet selectionnés
    long size = getSsAllObject( ssHatch, "" );
    
    //Verifier le nombre d'objets
    if( size != 0 )
    {
    
        //Nombre d'objet sélectionné
        int obj = 0;
        
        //Demander au dessinateur le chemin vers ou on enregistre le fichier
        file = askForFilePath( false,
                "xlsx;shp;csv;txt",
                "Enregistrer sous",
                current_folder );
                
        //Gestion d'erreur
        if( file == _T( "" ) )
        {
            print( "Commande annulée." );
            acedSSFree( ssHatch );
            return;
        }
        
        //Recuperer l'extension du fichier
        ext = getFileExt( file );
        
        //Changer l'encodage du chemin vers le fichier excel
        auto filename = latin1_to_utf8( acStrToStr( file ) );
        
        //Si fichier: .txt, .csv
        if( ext == "csv" || ext == "txt" )
        {
            //Ouvrir le fichier
            fstream fichier( acStrToStr( file ), ios::app );
            
            //Ecrire les en-têtes du fichier
            fichier << "Handle;Calque;Couleur;Type de ligne;Echelle de type de ligne;Epaisseur de ligne;Transparence;Materiau;X;Y;Z;Angle d'inclinaison;Surface;Motif" << endl;
            
            //Déclarer un AcDbHatch
            AcDbHatch* hatch = NULL;
            
            //Declarer un AcDbObjectId
            AcDbObjectId objId = NULL;
            
            //Barre de progression
            ProgressBar prog = ProgressBar( _T( "Progression:" ), size );
            
            //Recuperer les informations sur le fichier
            for( long i = 0; i < size; i++ )
            {
                //Declarer un AcDbEntity
                AcDbEntity* obje = NULL;
                
                //Recuperer le i-eme objid dans la selection
                objId = getObIdFromSs( ssHatch, i );
                
                //Verifier sa validité
                if( !objId.isValid() )
                    continue;
                    
                //Ouvrir l'objet
                acdbOpenAcDbEntity( obje, objId, AcDb::kForRead );
                
                //Verifier l'entité
                if( !obje )
                    continue;
                    
                //Determiner le type de l'AcDbObject
                if( obje->isKindOf( AcDbHatch::desc() ) )
                {
                    hatch = AcDbHatch::cast( obje );
                    
                    //Verifier l'entité
                    if( !hatch )
                    {
                        obje->close();
                        continue;
                    }
                    
                    //Recuperer les données du Hachures
                    exportHatch( hatch, pData );
                    
                    // Iteration sur le tableau des données
                    for( long j = 0; j < pData.size(); j++ )
                    {
                        //Ecriture des données dans le fichier sélectionné
                        fichier << pData[j] << ";";
                    }
                    
                    // Se mettre à la ligne
                    fichier << endl;
                    
                    //Incrémentation du nombre d'objet
                    obj++;
                    
                    // Vider le tableau des données
                    pData.clear();
                    
                    //Fermeture du texte
                    hatch->close();
                    prog.moveUp( i );
                }
                
                //Fermeture de l'objet
                obje->close();
                prog.moveUp( i );
            }
            
            try
            {
                //Fermeture du fichier
                fichier.close();
                
                //Affichage dans la console que les informations sont bien exportés
                print( "Exportation de " + to_string( obj ) + " hachures dans " + acStrToStr( file ) + " terminée." );
                
            }
            
            catch( ... )
            {
                print( "Impossible d'enregistrer le fichier." );
                acedSSFree( ssHatch );
                return;
            }
        }
        
        //Fichier: Excel
        else if( ext == "xlsx" )
        {
            // Initialisation du fichier excel
            workbook wb;
            worksheet sheet1 = wb.active_sheet();
            
            // Ecriture des en-têtes du fichier excel
            sheet1.cell( 1, 1 ).value( "Handle" );
            sheet1.cell( 2, 1 ).value( "Calque" );
            sheet1.cell( 3, 1 ).value( "Couleur" );
            sheet1.cell( 4, 1 ).value( "Type de ligne" );
            sheet1.cell( 5, 1 ).value( "Echelle de type de ligne" );
            sheet1.cell( 6, 1 ).value( "Epaisseur de ligne" );
            sheet1.cell( 7, 1 ).value( "Transparence" );
            sheet1.cell( 8, 1 ).value( "Materiau" );
            sheet1.cell( 9, 1 ).value( "X" );
            sheet1.cell( 10, 1 ).value( "Y" );
            sheet1.cell( 11, 1 ).value( "Z" );
            sheet1.cell( 12, 1 ).value( "Angle d'inclinaison" );
            sheet1.cell( 13, 1 ).value( "Surface" );
            sheet1.cell( 14, 1 ).value( "Motif" );
            
            //L'entité hatch
            AcDbHatch* hatch;
            AcDbObjectId objId = NULL;
            
            ProgressBar prog_xls = ProgressBar( _T( "Progression:" ), size );
            
            //3 - Récupérer les informations sur le/les hachure(s)
            int row = 2;
            
            for( long i = 0; i < size; i++ )
            {
            
                AcDbEntity* obje = NULL;
                //Le pointeur sur le texte
                objId = getObIdFromSs( ssHatch, i );
                
                //Safeguard
                if( !objId.isValid() )
                    continue;
                    
                //Ouvrir l'objet
                acdbOpenAcDbEntity( obje, objId, AcDb::kForRead );
                
                //Verifier l'entité
                if( !obje )
                    continue;
                    
                if( obje->isKindOf( AcDbHatch::desc() ) )
                {
                    hatch = AcDbHatch::cast( obje );
                    
                    //Verifier l'entité
                    if( !hatch )
                    {
                        obje->close();
                        continue;
                    }
                    
                    //Recuperer les données du Hachures
                    exportHatch( hatch, pData );
                    
                    // Mettre le format des cellules en texte
                    sheet1.columns( true ).number_format( number_format::text() );
                    
                    // Iteration sur les données
                    for( long k = 0; k < pData.size(); k++ )
                    {
                        //Ecriture dans le fichier excel
                        sheet1.cell( k + 1, row ).value( pData[k] );
                    }
                    
                    //Incrémentation du nombre d'objet
                    obj++;
                    
                    //Incrementer la ligne
                    row++;
                    
                    //Reinitialiser pData
                    pData.clear();
                    
                    //Fermer le hachure
                    hatch->close();
                    
                    //Progresser
                    prog_xls.moveUp( i );
                }
                
                //Fermeture de l'objet
                obje->close();
                prog_xls.moveUp( i );
            }
            
            try
            {
                // Enregistrer le fichier excel
                wb.save( filename );
                
                //Affichage dans la console que les informations sont bien exportés
                print( "Exportation de " + to_string( obj ) + " hachures dans " + acStrToStr( file ) + " terminée." );
                
                //Réinitialisation de la valeur de obj
                obj = 0;
            }
            
            catch( ... )
            {
                print( "Impossible d'enregistrer le fichier." );
                acedSSFree( ssHatch );
                return;
            }
        }
        
        //Fichier: Shape
        else if( ext == "shp" )
        {
            //Declarer une entité hatch
            AcDbHatch* hatch = NULL;
            writePrjOrNot( file );
            
            //Lire le fichier de parametre
            vector<AcString> field = createHatchField();
            
            //Creer un fichier shp
            SHPHandle myShapeFile = SHPCreate( file, SHPT_POINTZ );
            
            //Creer un dbfFile
            DBFHandle dbfHandle = DBFCreate( file );
            
            //Creer les champs
            int fieldNumber = createFields( dbfHandle, field );
            
            //Barre de progression
            ProgressBar prog = ProgressBar( _T( "Progression :" ), size );
            
            //Nombre d'élements exportés
            int nb_el = 0;
            
            AcDbObjectId objId;
            
            //Boucler sur toute les textes
            for( int i = 0; i < size; i++ )
            {
                //Declarer un AcDbEntity
                AcDbEntity* obje = NULL;
                
                //Recuperer le i-eme objid dans la selection
                objId = getObIdFromSs( ssHatch, i );
                
                //Verifier sa validité
                if( !objId.isValid() )
                    continue;
                    
                //Ouvrir l'objet
                acdbOpenAcDbEntity( obje, objId, AcDb::kForRead );
                
                //Verifier l'entité
                if( !obje )
                    continue;
                    
                //Determiner le type de l'AcDbObject
                if( obje->isKindOf( AcDbHatch::desc() ) )
                {
                    hatch = AcDbHatch::cast( obje );
                    
                    //Verifier l'entité
                    if( !hatch )
                    {
                        obje->close();
                        continue;
                    }
                    
                    //Exporter les hachures
                    drawHatchShp( hatch, myShapeFile, dbfHandle, field );
                    
                    //Incrementer le nombre d'element
                    nb_el++;
                }
                
                //Progresser
                prog.moveUp( i );
                
                //Liberer la memoire
                obje->close();
            }
            
            //Liberer la mémoire
            DBFClose( dbfHandle );
            SHPClose( myShapeFile );
            
            //Message
            print( "Exportation de " + to_string( nb_el ) + " hachures(s) dans " + acStrToStr( file ) + " terminée." );
        }
    }
    
    else
        print( "Aucune Hachure sélectionnée." );
        
    acedSSFree( ssHatch );
}

void cmdExportAtt()
{
    //Définition de la selection
    ads_name ssBlock;
    
    //Fichier & Extension
    AcString file, ext;
    
    file = _T( "" );
    
    //Recuperer le  chemin du dwg courant
    AcString current_folder = getCurrentFileFolder();
    current_folder += _T( "\\" );
    
    //Stockage : propriétés & données
    AcStringArray propName;
    
    //Vecteur temporaire: liste des attributs & liste des noms de propriete
    vector<AcString> nomAttributs;
    
    //Vecteur de type non fixe pour contenir les valeurs des proprietes
    vector<void*> propValue;
    
    //Set : (Noms des attributs sans doublons & Noms des propriétes sans doublons & Noms de définition) pour toutes les entités
    set<AcString> sAttr;
    
    //Nombre de blocs dynamiques
    int dynBlocCount = 0;
    
    //Message
    print( "Veuillez sélectionner le/les block(s) : " );
    
    //Recuperer le nombre de blockes
    long size = getSsBlock( ssBlock, "", "" );
    
    //Verifier
    if( size != 0 )
    {
    
        //Nombre d'objet exporté
        int obj = 0;
        
        //Demander chemin du fichier vers quoi on exporte
        file = askForFilePath( false,
                "csv;txt",
                "Enregistrer sous",
                current_folder );
                
        //Gestion d'erreur
        if( file == _T( "" ) )
        {
            print( "Commande annulée." );
            acedSSFree( ssBlock );
            return;
        }
        
        //Recuperer l'extension
        ext = getFileExt( file );
        
        //Changer l'encodage du chemin vers le fichier excel en utf8
        auto filename = latin1_to_utf8( acStrToStr( file ) );
        
        //Déclarer une reference de blocke
        AcDbBlockReference* blk;
        
        //Fichier : .txt ou .csv
        if( ext == "txt" || ext == "csv" )
        {
            //Ouvrir le fichier
            fstream fichier( acStrToStr( file ), ios::app );
            
            //Itération sur les objets
            for( long i = 0; i < size; i++ )
            {
                //Pointeur sur le bloc
                blk = getBlockFromSs( ssBlock, i );
                
                //Recuperer la taille
                int taille_attr = getAttributesNamesListOfBlockRef( blk ).size();
                
                if( taille_attr == 0 )
                    continue;
                    
                //Iteration sur la taille de la liste des noms des attributs
                for( long j = 0; j < taille_attr; j++ )
                {
                    //Mettre les attributs dans le set
                    sAttr.insert( getAttributesNamesListOfBlockRef( blk )[j] );
                }
                
                
                //Fermeture du bloc
                blk->close();
            }
            
            //Itération sur les données sans doublons dans le |setAttr| et les mettre dans le vecteur temporaire
            for( set<AcString>::iterator t = sAttr.begin(); t != sAttr.end(); t++ )
            {
                nomAttributs.push_back( *t );
                //Completer les en-tête avec la liste sans doublons des ATTRIBUTS
                fichier << acStrToStr( *t ) << ";";
            }
            
            //Se mettre à la ligne pour commencer à inserer les données
            fichier << endl;
            
            //Barre de progression
            ProgressBar prog = ProgressBar( _T( "Progression :" ), size );
            
            //Reiterer sur les blocks
            for( long i = 0; i < size; i++ )
            {
            
                //Pointeur sur le block
                blk = getBlockFromSs( ssBlock, i );
                
                //Verifier blk
                if( !blk )
                    continue;
                    
                //Recuperer la taille
                int taille_attr = getAttributesNamesListOfBlockRef( blk ).size();
                
                if( taille_attr == 0 )
                    continue;
                    
                //Rcuperer la taille de nomAttributs
                int nb_att = nomAttributs.size();
                
                //Itération sur les attributs
                for( long att = 0; att < nb_att; att++ )
                {
                    // Si la valeur de l'attribut n'est pas vide
                    if( getAttributValue( blk, nomAttributs[att] ) != "empty" )
                    {
                        //Convertir en std::string
                        auto attrib_val = acStrToStr( getAttributValue( blk, nomAttributs[att] ) );
                        
                        //Ecrire la valeur de l'attribut
                        fichier << attrib_val << ";";
                    }
                    
                    // Si la valeur de l'attribut est vide
                    else
                    {
                        // Remplir la colonne par des N/A
                        fichier << "N/A;";
                    }
                }
                
                //Se mettre à la ligne après une écriture d'une ligne d'information des objets
                fichier << endl;
                
                //Incrémentation de obj
                obj++;
                
                //Fermeture du block
                blk->close();
                
                //Progresser
                prog.moveUp( i );
            }
            
            try
            {
                //Fermeture du fichier
                fichier.close();
                
                //Affichage dans la console que les informations sont bien exportés
                print( "Exportation d'attributs de " + to_string( obj ) + " block(s) dans " + acStrToStr( file ) + " terminée." );
                
                //Réinitialisation de obj
                obj = 0;
            }
            
            catch( ... )
            {
                print( "Impossible d'enregistrer le fichier." );
                acedSSFree( ssBlock );
                return;
            }
            
        }
    }
    
    // Si aucun bloc n'est selectionné
    else
        print( "Aucun block sélectionné" );
        
    //Liberer la sélection
    acedSSFree( ssBlock );
}

void cmdExportRGBColor()
{

    //Recuperer le  chemin du dwg courant
    AcString current_folder = getCurrentFileFolder();
    current_folder += _T( "\\" );
    
    
    // test
    AcString file = _T( "" );
    
    // Demander à l'utilisateur le repertoire et l'extension du fichier à enregistrer
    file = askForFilePath( false,
            "xlsx",
            "Enregistrer sous",
            current_folder );
            
    // verifier file
    if( file == _T( "" ) )
        return;
        
        
    //Changer l'encodage du chemin pour le fichier excel
    auto filename = latin1_to_utf8( acStrToStr( file ) );
    
    // créer les couleurs
    map<int, vector<int>> colors;
    
    // exporter les couleurs
    for( int k = 0; k < 256; k++ )
    {
    
        // créer un objet accmcolor
        AcCmColor col;
        col.setColorIndex( k );
        
        
        // recupere les indices rgb
        colors.insert( { k, {col.red(), col.green(), col.blue()} } );
    }
    
    
    // Initialisation du fichier excel
    workbook wb;
    worksheet ws = wb.active_sheet();
    ws.title( "RGBCOLOR" );
    
    // Ajout des en-têtes dans le feuille excel
    ws.cell( 1, 1 ).value( "Indice" );
    ws.cell( 2, 1 ).value( "R" );
    ws.cell( 3, 1 ).value( "G" );
    ws.cell( 4, 1 ).value( "B" );
    
    // Mettre le format des cellules en texte
    ws.columns( true ).number_format( number_format::text() );
    
    // compteur des lignes
    int rowCounter = 2;
    
    for( auto col : colors )
    {
    
        // écriture dans le fichier excel
        ws.cell( 1, rowCounter ).value( col.first );
        ws.cell( 2, rowCounter ).value( col.second[0] );
        ws.cell( 3, rowCounter ).value( col.second[1] );
        ws.cell( 4, rowCounter ).value( col.second[2] );
        
        // incrementer le compteur de ligne
        rowCounter++;
    }
    
    try
    {
        // Enregistrer le fichier excel
        wb.save( filename );
        
        //Affichage dans la console que les informations sont bien exportés
        print( "Exportation des couleurs dans " + acStrToStr( file ) + " terminée." );
        
    }
    
    catch( ... )
    {
        print( "Impossible d'enregistrer le fichier." );
        return;
    }
    
}