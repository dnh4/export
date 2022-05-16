#include "acrxEntryPoint.h"


void unloadApp()
{
    //Supprimer le groupe
    delFmapCmd( sGroupName );
}


extern "C" AcRx::AppRetCode
acrxEntryPoint( AcRx::AppMsgCode msg, void* pkt )
{
    switch( msg )
    {
        //Chargement de l'application
        case AcRx::kInitAppMsg:
            acrxDynamicLinker->unlockApplication( pkt );
            acrxRegisterAppMDIAware( pkt );
            initApp();
            break;
            
        //Dechargement de l'application
        case AcRx::kUnloadAppMsg:
            unloadApp();
            break;
            
        //Autre
        default:
            break;
            
    }
    
    return AcRx::kRetOK;
}

void initApp()
{
    //Export polylignes et lignes
    addFmapCmd( cmdExportPoly3d, "EXPORTPOLY3D", sGroupName ); //Ok
    addFmapCmd( cmdExportPoly2d, "EXPORTPOLY2D", sGroupName ); //Ok
    addFmapCmd( cmdExportPoly, "EXPORTPOLY", sGroupName ); //Ok
    addFmapCmd( cmdExportClosedPoly2d, "EXPORTCLOSEDPOLY2D", sGroupName ); //Ok
    addFmapCmd( cmdExportClosedPoly3d, "EXPORTCLOSEDPOLY3D", sGroupName ); //Ok
    addFmapCmd( cmdExportClosedPoly, "EXPORTCLOSEDPOLY", sGroupName ); //Ok
    addFmapCmd( cmdExportLine, "EXPORTLINE", sGroupName ); //Ok
    
    //Export block
    addFmapCmd( cmdExportBlock, "EXPORTBLOCK", sGroupName ); //Ok
    addFmapCmd( cmdExportBlockPoly, "EXPORTBLOCKPOLY", sGroupName ); //Ok
    addFmapCmd( cmdExportBlockPoly2D, "EXPORTBLOCKPOLY2D", sGroupName ); //Ok
    addFmapCmd( cmdExportBlockPoly3D, "EXPORTBLOCKPOLY3D", sGroupName ); //Ok
    
    //Export textes
    addFmapCmd( cmdExportTextPoly, "EXPORTTEXTPOLY", sGroupName ); //Ok
    addFmapCmd( cmdExportText, "EXPORTTEXT", sGroupName ); //Ok
    addFmapCmd( cmdExportMText, "EXPORTMTEXT", sGroupName ); //Ok
    addFmapCmd( cmdExportTextPoly2D, "EXPORTTEXTPOLY2D", sGroupName ); //Ok
    addFmapCmd( cmdExportTextPoly3D, "EXPORTTEXTPOLY3D", sGroupName ); //Ok
    
    //Export curves
    addFmapCmd( cmdExportCurve, "EXPORTCURVE", sGroupName ); //Ok
    
    //Export points
    addFmapCmd( cmdExportPoint, "EXPORTPOINT", sGroupName ); //Ok
    addFmapCmd( cmdExportPonctuel, "EXPORTPONCTUEL", sGroupName ); //Ok
    
    //Export to XML
    addFmapCmd( cmdExportPoly3DToXml, "POLY3DTOXML", sGroupName );
    addFmapCmd( cmdMoinsFaceToXml, "-FACETOXML", sGroupName );
    addFmapCmd( cmdFaceToXml, "FACETOXML", sGroupName );
    
    //Export entity
    addFmapCmd( cmdExportEntity, "EXPORTENTITY", sGroupName ); //Ok
    
    //Export Circle
    addFmapCmd( cmdExportCircle, "EXPORTCIRCLE", sGroupName ); //Ok
    
    //Export Hatch
    addFmapCmd( cmdExportHatch, "EXPORTHATCH", sGroupName ); //Ok
    
    //Export Trait
    addFmapCmd( cmdExportAtt, "EXPORTATT", sGroupName );
    
    // Export color
    addFmapCmd( cmdExportRGBColor, "EXPORTRGBCOLOR", sGroupName );
}
