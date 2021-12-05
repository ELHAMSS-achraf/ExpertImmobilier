import { Component, OnInit } from '@angular/core';
import * as fs from "fs";
import { FormControl } from '@angular/forms';
import { saveAs } from "file-saver";
import { Document, Table, TableRow, WidthType, BorderStyle, TableCell, AlignmentType, Packer, Paragraph, TextRun, HeadingLevel, ShadingType, SectionType, ImageRun } from "docx";

@Component({
  selector: 'app-generator-report',
  templateUrl: './generator-report.component.html',
  styleUrls: ['./generator-report.component.css']
})
export class GeneratorReportComponent {

  constructor() { }

  //     let rawdata = fs.readFileSync('Data.json').toString("utf8");
  // let expert = JSON.parse(rawdata);
  
  // Documents contain sections, you can have multiple sections per document, go here to learn more about sections
  // This simple example will only contain one section
  // public getVal(val) : string{
  //     var a;
  //     a= console.warn(val);
  //     return a;
  // };


  public onSubmit(id) : string{
  var b= "";
  b = (<HTMLInputElement>document.getElementById(id)).value;
  return b}

  public onSubmit1(id) : string{
    var h= "";
    h = (<HTMLTextAreaElement>document.getElementById(id)).value;
    return h}
    public nature(id1, id2, id3, id4) :string { 
        var a =document.getElementById(id1) as HTMLInputElement;
        var b =document.getElementById(id1) as HTMLInputElement;
        var c =document.getElementById(id1) as HTMLInputElement;
        var d =document.getElementById(id1) as HTMLInputElement;
        var e="";

        if(a != null && a.checked) {
            e = "Terrain";
            return e;
        }
        else if(b != null && b.checked) {
            e = "Bureau";
            return e;
        }
        else if(c != null && c.checked) {
            e = "Immeuble";
            return e;
        }
        else if(d != null && d.checked) {
            e = "Commerce";
            return e;
        }
        else{
            return "";
        }


    }


  public display1(id3) :string { 
    var t =document.getElementById(id3) as HTMLInputElement;
    if(t != null && t.checked) {
        var c = "X";
        return c
    }
    else{
        return "";
    }
}
public display2(id2) :string { 
    var t =document.getElementById(id2) as HTMLInputElement;
    if(t != null && t.checked) {
        var c = "X";
        return c
    }
    else{
        return "";
    }
}
  public download(): void {
    const table = new Table({
        width: {
            size: 9070,
            type: WidthType.DXA,
        },
        rows: [
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("      ")],
                    }),
                    new TableCell({
                        children: [new Paragraph("OUI")],
                    }),
                    new TableCell({
                        children: [new Paragraph("NON")],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Plein centre")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1("OUI"))],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.display2("NON"))
                        ],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Proche centre (300)")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1("centre"))],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.display2("centre2"))
                        ],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Périphérie")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1("per"))],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.display2("per2"))
                        ],
                    }),
                ],
            }),


            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Terrain sur Rue")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1("rue"))],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.display2("rue2"))
                        ],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Mitoyen ")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1("Mitoyen"))],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.display2("Mitoyen2"))
                        ],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Vis à Vis ")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1("vis"))],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.display2("vis2"))
                        ],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Espaces verts ")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1("vert"))],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.display2("vert2"))
                        ],
                    }),
                ],
            }),
            


        ],
    });



    const table7 = new Table({
        width: {
            size: 9070,
            type: WidthType.DXA,
        },
        rows: [
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("  ")],
                    }),
                    new TableCell({
                        children: [new Paragraph("OUI ")],
                    }),
                    new TableCell({
                        children: [new Paragraph("Non ")],
                    }),
                    new TableCell({
                        children: [new Paragraph("OBSERVATIONS ")],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Remblais  ")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1('Remblais'))],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.display2('Remblais2'))
                        ],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph("    ")
                        ],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Présence de cuves   ")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1('cuve'))],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.display2('cuve2'))
                        ],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph("    ")
                        ],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Anciennes activité poulluantes   ")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1('activite'))],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display2('activite2'))],
                    }),
                    new TableCell({
                        children: [new Paragraph("     ")],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("prescriptions archéo  ")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1("archeo"))],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display2('archeo2'))],
                    }),
                    new TableCell({
                        children: [new Paragraph("OBSERVATIONS ")],
                    }),
                ],
            }),
    
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Présence de transformateur  ")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1('trans'))],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display2('trans2'))],
                    }),
                    new TableCell({
                        children: [new Paragraph("  ")],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Inondation  ")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1('ino'))],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display2('ino2'))],
                    }),
                    new TableCell({
                        children: [new Paragraph("   ")],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Reprises en sous oeuvres  ")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1('ouvre'))],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display2('ouvre2'))],
                    }),
                    new TableCell({
                        children: [new Paragraph("  ")],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Présence d'une voie ferré  ")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1('voi'))],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display2('voi2'))],
                    }),
                    new TableCell({
                        children: [new Paragraph("    ")],
                    }),
                ],
            }),

            
            
            

        ],
    });



    const table5 = new Table({
        width: {
            size: 9070,
            type: WidthType.DXA,
        },
        rows: [
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("      ")],
                    }),
                    new TableCell({
                        children: [new Paragraph("OUI")],
                    }),
                    new TableCell({
                        children: [new Paragraph("NON")],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Sonores")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1("son"))],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.display2("NON"))
                        ],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Proche centre (300)")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1("centre"))],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.display2("centre2"))
                        ],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Périphérie")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1("per"))],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.display2("per2"))
                        ],
                    }),
                ],
            }),


            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Terrain sur Rue")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1("rue"))],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.display2("rue2"))
                        ],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Mitoyen ")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1("Mitoyen"))],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.display2("Mitoyen2"))
                        ],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Vis à Vis ")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1("vis"))],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.display2("vis2"))
                        ],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Espaces verts ")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1("vert"))],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.display2("vert2"))
                        ],
                    }),
                ],
            }),
            


        ],
    });

    const nuissanceTable = new Table({
        width: {
            size: 9070,
            type: WidthType.DXA,
        },
        rows: [
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("      ")],
                    }),
                    new TableCell({
                        children: [new Paragraph("OUI")],
                    }),
                    new TableCell({
                        children: [new Paragraph("NON")],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Sonores")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1("sono"))],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.display2("sono2"))
                        ],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Lignes à Haute tesion")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1("ligne"))],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.display2("ligne2"))
                        ],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("visuelles")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1("visu"))],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.display2("visu2"))
                        ],
                    }),
                ],
            }),


            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Olfactives")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1("fact"))],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.display2("fact2"))
                        ],
                    }),
                ],
            }),
            


        ],
    });
    

    const clientTable = new Table({
        width: {
            size: 9070,
            type: WidthType.DXA,
        },
        rows: [
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Client")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.onSubmit('name'))],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Objet de l'evaluation")],
                    }),
                    new TableCell({
                        children: [new Paragraph("Valeur vénale et marchande de l’actif "+ this.onSubmit('Pptedite'))],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Interlocuteur Client")],
                    }),
                    new TableCell({
                        children: [new Paragraph("Monsieur "+ this.onSubmit('name'))],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("PIÈCES RECUPEREES AUPRES DE LA CONSERVATION ")],
                        rowSpan: 3,
                    }),
                    new TableCell({
                        children: [new Paragraph("Documents")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph("Statut")
                        ],
                    }),
                    new TableCell({
                        children: [
                            new Paragraph("Date de Document")
                        ],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Certificat de proprièté")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.onSubmit('certif'))],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(" ")
                        ],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Status de Plan Cadastral")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.onSubmit('splan'))],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(" ")
                        ],
                    }),
                ],
            }),

        ],
    });
    

    const table8 = new Table({
        width: {
            size: 9070,
            type: WidthType.DXA,
        },
        rows: [
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Date de la visite")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.onSubmit('date'))],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Personnes presentées")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.onSubmit('visite'))],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Commentaires")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.onSubmit1('tr'))],
                    }),
                ],
            }),


        ],
    });
    





    const juridiqueTable = new Table({
        width: {
            size: 9070,
            type: WidthType.DXA,
        },
        rows: [
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Numéro de TF ")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.onSubmit('TF'))],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Intitulé de la propriété ")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit('Pptedite'))
                        ],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Consistance ")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph("  ")
                        ],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Conservation Foncière")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit('city'))
                        ],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Propriétaire")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit('name'))
                        ],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Superficie ")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit('surface'))
                        ],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Droit réel ou charge foncière")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit('reel'))
                        ],
                    }),
                ],
            }),


        ],
    });

    const table10 = new Table({
        width: {
            size: 9070,
            type: WidthType.DXA,
        },
        rows: [
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Périmétre de l'étude ")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.onSubmit1('perm'))],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Offre présente sur le périmètre étudié")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit1('etd'))
                        ],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Baromètre du marché ")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit1('baro'))
                        ],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Différenciation du bien par rapport à l’offre")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit1('offre'))
                        ],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Prix de vente ")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit1('prix'))
                        ],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Évolution & projets futurs de la zone")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit1('future'))
                        ],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Droit réel ou charge foncière")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit('reel'))
                        ],
                    }),
                ],
            }),


        ],
    });

    const table1 = new Table({
        width: {
            size: 9070,
            type: WidthType.DXA,
        },
        rows: [
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("NOM DU DOSSIER ")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.onSubmit('dossier'))],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Objet de La mission")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph("Évaluation en valeur vénale")
                        ],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Date d'expertise")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit('exp'))
                        ],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Adresse des actifs")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit('actif'))
                        ],
                    }),
                ],
            }),


        ],
    });

    const table9 = new Table({
        width: {
            size: 9070,
            type: WidthType.DXA,
        },
        rows: [
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Catégories ")],
                    }),
                    new TableCell({
                        children: [new Paragraph("Caractéristiques")],
                    }),
                    new TableCell({
                        children: [new Paragraph("Notes")],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Générale")],
                        rowSpan: 6,
                    }),
                    new TableCell({children: [new Paragraph("Orientation (bien traversant)") ],}),
                    new TableCell({children: [new Paragraph(this.onSubmit('or')) ],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Présence de place de parking")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit('park'))
                        ],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Etat des parties communes")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit('des'))
                        ],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Organisation spatiale")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit('spa'))
                        ],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Etat des revêtements")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit('rev'))
                        ],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("EAccessibilité du bien")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit('ben'))
                        ],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Technique")],
                        rowSpan: 6,
                    }),
                    new TableCell({
                        children: [new Paragraph("Etanchéité")],
                       
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit('etan'))
                        ],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Etat de la boiserie")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit('boi'))
                        ],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Etat des revêtements de sol")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit('sol'))
                        ],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Electricité")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit('do'))
                        ],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Totale")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph("   ")
                        ],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph("   ")
                        ],
                    }),
                ],
            }),



        ],
    });


    const table4 = new Table({
        width: {
            size: 9070,
            type: WidthType.DXA,
        },
        rows: [
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("          ")],
                    }),
                    new TableCell({
                        children: [new Paragraph("DISTANCE  ")],
                    }),
                    new TableCell({
                        children: [new Paragraph("OBSERVATIONS  ")],
                    }),
                ],
            }),


        ],
    });
//    const image = new ImageRun({
//     data: fs.readFileSync("/src/assets/img/Capture.PNG"),
//     transformation: {
//         width: 200,
//         height: 200,
//     }
// });

    const table3 = new Table({
        width: {
            size: 9070,
            type: WidthType.DXA,
        },
        rows: [
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Commerces ")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.onSubmit1('commerce'))],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Services ")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit1('service'))
                        ],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Ecole")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit1('ecole'))
                        ],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Transport")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit1('trsp'))
                        ],
                    }),
                ],
            }),


        ],
    });



    const table2 = new Table({
        width: {
            size: 9070,
            type: WidthType.DXA,
        },
        rows: [
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Commune ")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.onSubmit('commune'))],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Surface globale des biens (m²)")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit("surface"))
                        ],
                    }),
                ],
            }),
            

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Numéro de titre foncier")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit('TF'))
                        ],
                    }),
                ],
            }),


        ],
    });



    const situationTable = new Table({
        width: {
            size: 9070,
            type: WidthType.DXA,
        },
        rows: [
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Situation ")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.onSubmit('sit'))],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Objet d'usage  ")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit("obj"))
                        ],
                    }),
                ],
            }),
            

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Caractère de la zone ")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit('zone'))
                        ],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Evolution & Projection de la zone ")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit1('prj'))
                        ],
                    }),
                ],
            }),

        ],
    });
    
    

    const proxTable = new Table({
        width: {
            size: 9070,
            type: WidthType.DXA,
        },
        rows: [
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Commerces ")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.onSubmit1('cmr'))],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Administrations publics  ")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit1("adm"))
                        ],
                    }),
                ],
            }),
            

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Ecoles  ")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit1('col'))
                        ],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Agences de services ")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit1('srv'))
                        ],
                    }),
                ],
            }),

        ],
    });
    


    
  const doc = new Document({
      
      sections: [{
          properties: {
              type: SectionType.NEXT_PAGE,
          },
          children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                heading: HeadingLevel.TITLE,
                children: [
                    new TextRun({
                  text: "RAPPORT D’EXPERTISE",
                  bold: true,
                  italics: true,
                  color: '000000',
                  allCaps: true,
                  size: 50,
                  break: 1,
                  

                    }),
                    new TextRun({
                        text : "Proprièté dite "+ this.onSubmit('Pptedite'),
                        color : '000000',
                        size : 32,
                        break: 1
                    }),
                    new TextRun({
                        text : "Objet de titre foncier N°  "+ this.onSubmit('TF'),
                        color : '000000',
                        size : 32,
                        break: 1
                    }),
                    new TextRun({
                        text : "Situé à :  "+ this.onSubmit('Pptedite'),
                        color : '000000',
                        size : 32,
                        break: 1
                    }),
                  
                    // new ImageRun({
                    //     data:  readFileSync(this.onSubmit('image')) ,
                    //     transformation: {
                    //         width: 200,
                    //         height: 200,
                    //     },
                    // }),
                    
                    

                ],
                  spacing: {
                      before: 250,
                      after: 9000,
                  },
                  
              }),
              
              new Paragraph({
                  
                          text:"Sommaire",
                          
                          heading: HeadingLevel.TITLE,
                         
                         
                          
                          shading: {
                              type: ShadingType.REVERSE_DIAGONAL_STRIPE,
                              color: "4682B4",
                              fill: "4682B4",
                          },
                          spacing: {
                              after: 200
                          }
                      }),
              new Paragraph({
                          text: "\t1-1.	La mission",
                        
                          heading: HeadingLevel.HEADING_1,
                         
                          spacing: {
                              after: 200
                          }
                      }),
              new Paragraph({
                          text: "\t1-2.	Certificat d’évaluation",
                         
                          heading: HeadingLevel.HEADING_1,
                          
                          spacing: {
                              after: 200
                          }
                      }),
              new Paragraph({
                          text: "\t1-3	DEFINITIONS",
                        
                          heading: HeadingLevel.HEADING_1,
                          
                          spacing: {
                              after: 500
                              
                          }
                      }),
                      new Paragraph({
                          text: "\t2-1.	DEFINITION DU BIEN",
                          
                          heading: HeadingLevel.HEADING_1,
                          
                          spacing: {
                              after: 200
                          }
                      }),
                      new Paragraph({
                          text: "\t2-2.	VISITE DU BIEN",
                         
                          heading: HeadingLevel.HEADING_1,
                          
                          spacing: {
                              after: 200
                          }
                      }),
              new Paragraph({
                          text: "\t2-3.	DESCRIPTION DU BIEN",
                          
                          heading: HeadingLevel.HEADING_1,
                          
                          spacing: {
                              after: 500
                              
                          }
                      }),
  
                      new Paragraph({
                          text: "\t3-1.	ETAT APPARENT DU BIEN",
                       
                          heading: HeadingLevel.HEADING_1,
                         
                          spacing: {
                              after: 200
                          }
                      }),
                      new Paragraph({
                          text: "\t3-2.	SITUATION FONCIER – JURIDIQUE",
                         
                          heading: HeadingLevel.HEADING_1,
                          
                          spacing: {
                              after: 200
                          }
                      }),
                      new Paragraph({
                          text: "\t3-3.	SITUATION URBANISTIQUE",
                         
                          heading: HeadingLevel.HEADING_1,
                          
                          spacing: {
                              after: 200
                          }
                      }),
              new Paragraph({
                          text: "\t3-4.	ETUDE DE MARCHE",
                       
                          heading: HeadingLevel.HEADING_1,
                          
                          spacing: {
                              after: 9000
                              
                          }
                      }),
                      new Paragraph({
                          text: "1. La Mission :   ",
                          
                          heading: HeadingLevel.HEADING_2,
                          
                          border: {
                              bottom: {
                                  color: "auto",
                                  space: 1,
                                  value: "single",
                                  size: 6,
                              },
                          },
                          spacing: {
                              after: 200
                              
                          }
                      }),
                      new Paragraph({
                          text: "Rapport d’évaluation du TF :"+ this.onSubmit('TF'),
                          heading: HeadingLevel.HEADING_2,
                         
                      
                          spacing: {
                              after: 200
                          }
                      }),
                      new Paragraph({
                        children: [
                            new TextRun({
                                text: "Agissant suite à la demande de Monsieur  " + this.onSubmit("name")+" "  + +" en qualité de "+ this.onSubmit('qualite') +", tendant à définir la valeur vénale des constructions de la propriété dite "+this.onSubmit('Pptedite') + " objet du titre foncier  n° "+ this.onSubmit('TF')+" sise : "+ this.onSubmit('adresse') + ". Nous procédons à une étude d’évaluation des actifs corporels (Terrain, bâtiments , constructions,  installations et autres) qui nous ont été présentés et vous présentons nos conclusions dans le présent rapport.",
                                bold: false,
                                color: "000000",
                            }),
                            new TextRun({
                                text: ". Nous procédons à une étude d’évaluation des actifs corporels (Terrain, bâtiments , constructions,  installations et autres) qui nous ont été présentés et vous présentons nos conclusions dans le présent rapport.",
                                bold: false,
                                color: "000000",
                            }),],
                       
                          heading: HeadingLevel.HEADING_3,
                          
                          spacing: {
                              after: 200
                              
                          }
                      }),
                      new Paragraph({
                        children: [
                            new TextRun({
                          text: "\t1-	Enquête préalable auprès du requérant, notamment pour recueillir des informations sur les titres de la propriété, plans existants, certificat de propriété, ainsi que des éléments comptables concernant les plantations, les constructions et les équipements.",
                          bold: false,
                          color: "000000",
                      }),],
                          heading: HeadingLevel.HEADING_3,
                          
                          spacing: {
                              after: 200
                          }
                      }),
                      new Paragraph({
                        children: [
                            new TextRun({
                          text: "\t2-	Visite des lieux en date du"+this.onSubmit("date") +" pour appréciation des équipements et constructions réellement existantes, date de mise en place, avec quelques métrés et enquête auprès des responsables sur place.",
                          bold: false,
                          color: "000000",
                      }),],
                          heading: HeadingLevel.HEADING_3,
                          
                          spacing: {
                              after: 200
                          }
                      }),
                      new Paragraph({
                        children: [
                            new TextRun({
                          text: "\t3-	Consultation des données informatiques (Google Maps) avec insertion des données topographiques (coordonnées Lambert). ",
                          bold: false,
                          color: "000000",
                      }),],
                          heading: HeadingLevel.HEADING_3,
                          
                          spacing: {
                              after: 200
                          }
                      }),
                      new Paragraph({
                        children: [
                            new TextRun({
                          text: "\t4-	Définition de chaque élément de la propriété aux fins d’en estimer les valeurs.",
                          bold: false,
                          color: "000000",
                      }),],
                          heading: HeadingLevel.HEADING_3,
                          
                          spacing: {
                              after: 200
                              
                          }
                      }),
                      new Paragraph({
                        children: [
                            new TextRun({
                          text: "\t5-	Enfin, nous avons mené une étude pour définir la valeur vénale de la propriété.",
                          bold: false,
                          color: "000000",
                      }),],
                          heading: HeadingLevel.HEADING_3,
                          // shading: {
                          //   type: ShadingType.REVERSE_DIAGONAL_STRIPE,
                          //   color: "000000",
                          //   fill: "FFFFFF"},
                          spacing: {
                              after: 200
                              
                          }
                      }),
                      new Paragraph({
                        children: [
                            new TextRun({
                          text: "\t6-	Nous utilisons la méthode de "+this.onSubmit('mtd1') +", la méthode de "+ this.onSubmit('mtd2')+" et la méthode de " + this.onSubmit('mtd3'),
                          bold: false,
                          color: "000000",
                      }),],
                          heading: HeadingLevel.HEADING_3,
                          
                          spacing: {
                              after: 9000
                              
                          }
                      }),
                      new Paragraph({
                        text: "2. Certificat d’évaluation :",
                        heading: HeadingLevel.HEADING_2,
                        border: {
                            bottom: {
                                color: "auto",
                                space: 1,
                                value: "single",
                                size: 6,
                            },
                        },
                    
                        spacing: {
                            after: 200
                        }
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({
                                text:"A Monsieur "+ this.onSubmit("name"),
                          bold: false,
                          color: "000000",
                      }),],
                          heading: HeadingLevel.HEADING_3,
                          // shading: {
                          //   type: ShadingType.REVERSE_DIAGONAL_STRIPE,
                          //   color: "000000",
                          //   fill: "FFFFFF"},
                          spacing: {
                              after: 200
                              
                          }
                      }),
                      new Paragraph({
                        children: [
                            new TextRun({
                                text : "Objet : Estimation de la valeur d’un actif immobilier sis à "+ this.onSubmit("adresse"),
                          bold: true,
                          color: "000000",
                      }),],
                          heading: HeadingLevel.HEADING_3,
                          
                          spacing: {
                              after: 600
                              
                          }
                      }),
                      new Paragraph({
                        children: [
                            new TextRun({
                                text: "Monsieur,",
                          bold: false,
                          color: "000000",
                      }),]}),
                      new Paragraph({
                        children: [
                      new TextRun({
                        text: "Nous avons l’honneur de vous communiquer notre estimation de la valeur d’un actif immobilier objet du TF 244104/07 Au regard des éléments qui nous ont été communiqués, de notre évaluation et des analyses décrites dans notre présent rapport d’expertise, nous retiendrons une Valeur Vénale (VV) de 737000 dhs (var : valeur venale)en date du 06-Feb-21pour (var : date rapport)cet actif libre de toutes occupations, charges ou   contrainte,",
                  bold: false,
                  color: "000000",
              }),
                    ],
                          heading: HeadingLevel.HEADING_3,
                          
                          spacing: {
                              after: 600
                              
                          }
                      }),
                      new Paragraph({
                        children: [
                            new TextRun({
                                text: "Nous restons à votre entière disposition pour vous apporter toute information qui vous semblerait nécessaire sur le contenu du rapport d’expertise et la méthodologie adoptée.  \n",
                                bold: false,
                          color: "000000",
                      }),
                      new TextRun({
                          text: "Nous vous prions d’agréer, Monsieur, l’expression de nos salutations distinguées.",
                        bold: false,
                  color: "000000",
              }),
                    ],
                          heading: HeadingLevel.HEADING_3,
                          // shading: {
                          //   type: ShadingType.REVERSE_DIAGONAL_STRIPE,
                          //   color: "000000",
                          //   fill: "FFFFFF"},
                          spacing: {
                              after: 300
                              
                          }
                      }),
                      new Paragraph({
                        text: "Salim BENMLIH  : Ingénieur Expert Immobilier",
                        heading: HeadingLevel.HEADING_2,
                        spacing: {
                            after: 200
                        }
                    }),
                    new Paragraph({
                        text:"Mohammed Soussi Sadoq : Notaire Expert Immobilier RICS",
                        heading: HeadingLevel.HEADING_2,
                        spacing: {
                            after: 5000
                        }
                    }),
                    new Paragraph({
                        text :"3. Définitions :       ",
                        heading: HeadingLevel.HEADING_2,
                        border: {
                            bottom: {
                                color: "auto",
                                space: 1,
                                value: "single",
                                size: 6,
                            },
                        },
                    
                        spacing: {
                            after: 200
                        }
                    }),
                    new Paragraph({
                        text : "Evaluation (Valuation) : Avis écrit d’un expert sur la valeur des droits liés à un actif immobilier, à la date d’évaluation. L’évaluation sera établie, sauf limites convenues dans les conditions du contrat, à la suite d’une visite, et de toute investigation ou enquête complémentaire pertinente en conformité avec la nature de l’actif immobilier et les objectifs définis dans le contrat d’évaluation.  ",
                        heading: HeadingLevel.HEADING_3,
                        spacing: {
                            after: 200
                        }
                    }),
                    new Paragraph({
                        text : "Immeuble (Property) : Ensemble des droits et intérêts portant sur des terrains (avec et sans constructions), installations techniques et équipements et tout actif pouvant se déprécier sauf définition plus restrictive dictée par le contexte. Cette expression s’applique également aux autres actifs détenus tels que des stocks ou travaux en cours, lorsque l’évaluation a pour but d’intégrer une valeur dans les états financiers.  ",
                        heading: HeadingLevel.HEADING_3,
                        spacing: {
                            after: 200
                        }
                    }),
                    new Paragraph({
                        text: "Juste Valeur (Fair Value) : Montant pour lequel un actif peut être échangé ou une dette réglée entre des parties bien informées, consentantes et agissant dans des conditions de concurrence normales.  ",
                        heading: HeadingLevel.HEADING_3,
                        spacing: {
                            after: 200
                        }
                    }),
                    new Paragraph({
                        text: "Méthode d’évaluation (Method of Valuation) : Approche ou technique employée pour calculer la valeur définie selon le cadre d’évaluation.  ",
                        heading: HeadingLevel.HEADING_3,
                        spacing: {
                            after: 200
                        }
                    }),
                    new Paragraph({
                        text: "Valeur vénale (VV) Market Value :  Somme d’argent estimée contre laquelle un immeuble serait échangé à la date d’évaluation entre un acheteur consentant et un vendeur consentant dans une transaction équilibrée après une commercialisation adéquate où les parties ont l’une et l’autre agit en toute connaissance, prudemment et sans pression.     ",
                        heading: HeadingLevel.HEADING_3,
                        spacing: {
                            after: 200
                        }
                    }),
                    new Paragraph({
                        text: "Source : NORMES D’EVALUATION DE LA RICS (sixième édition, version Française, avril 2010)   ",
                        heading: HeadingLevel.HEADING_3,
                        spacing: {
                            after: 1900
                        }
                    }),
                    new Paragraph({
                  
                        children: [
                            new TextRun({
                                text: "       I - Le site ",
                          bold: false,
                          color: "000000",
                      }),],
                        
                        heading: HeadingLevel.HEADING_1,
                       
                       
                        
                        shading: {
                            type: ShadingType.REVERSE_DIAGONAL_STRIPE,
                            color: "4682B4",
                            fill: "4682B4",
                        },
                        spacing: {
                            after: 200
                        }
                    }),
                    table1,
                    table2,
                    new Paragraph({
                  
                        
                        children: [
                            new TextRun({
                                text:"     SITUATION               ACCESSIBILITE (Distance en m2)",
                                bold: false,
                          color: "000000",
                      }),],
                        
                        heading: HeadingLevel.HEADING_1,
                       
                        
                        shading: {
                            type: ShadingType.REVERSE_DIAGONAL_STRIPE,
                            color: "4682B4",
                            fill: "4682B4",
                        },
                        spacing: {
                            after: 200
                        }
                    }),
                    table,
                    new Paragraph({
                  
                        children: [
                            new TextRun({
                                text: "       Observations ",
                          bold: false,
                          color: "000000",
                      }),],
                        
                        heading: HeadingLevel.HEADING_1,
                       
                       
                        
                        shading: {
                            type: ShadingType.REVERSE_DIAGONAL_STRIPE,
                            color: "4682B4",
                            fill: "4682B4",
                        },
                        spacing: {
                            after: 200
                        }
                    }),
                    
                    table3,

                    new Paragraph({
                  
                        
                        children: [
                            new TextRun({
                                text:"     DESSERTE",
                                bold: true,
                          color: "000000",
                      }),],
                        
                        heading: HeadingLevel.HEADING_1,
                       
                        
                        shading: {
                            type: ShadingType.REVERSE_DIAGONAL_STRIPE,
                            color: "4682B4",
                            fill: "4682B4",
                        },
                        spacing: {
                            after: 200
                        }
                    }),
                    table4,

                    new Paragraph({
                  
                        
                        children: [
                            new TextRun({
                                text:"     NUISSANCES",
                                bold: true,
                          color: "000000",
                      }),],
                        
                        heading: HeadingLevel.HEADING_1,
                       
                        
                        shading: {
                            type: ShadingType.REVERSE_DIAGONAL_STRIPE,
                            color: "4682B4",
                            fill: "4682B4",
                        },
                        spacing: {
                            after: 200
                        }
                    }),
                    nuissanceTable,
                    new Paragraph({
                  
                        
                        children: [
                            new TextRun({
                                text:"     Risques",
                                bold: true,
                          color: "000000",
                      }),],
                        
                        heading: HeadingLevel.HEADING_1,
                       
                        
                        shading: {
                            type: ShadingType.REVERSE_DIAGONAL_STRIPE,
                            color: "4682B4",
                            fill: "4682B4",
                        },
                        spacing: {
                            after: 200,
                            before : 20
                        }
                    }),
                    table7,
                    new Paragraph({children: [
                            new TextRun({
                                text:"     Observations Particulières",
                                bold: true,
                          color: "000000",
                      }),],
                        heading: HeadingLevel.HEADING_1,
                        shading: {type: ShadingType.REVERSE_DIAGONAL_STRIPE,
                            color: "4682B4",
                            fill: "4682B4",},
                        spacing: {after: 200,
                            before : 20}}),
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: this.onSubmit1('obs'),
                          bold: false,
                          color: "000000",
                      }),],
                      spacing: {
                        after: 200,
                        before : 20
                    }
                    }),
                    new Paragraph({children: [
                        new TextRun({
                            text:"     II - METHODOLOGIE D'EVALUATION ",
                            bold: true,
                      color: "000000",
                  }),],
                    heading: HeadingLevel.HEADING_1,
                    shading: {type: ShadingType.REVERSE_DIAGONAL_STRIPE,
                        color: "4682B4",
                        fill: "4682B4",},
                    spacing: {after: 200,
                        before : 20}}),
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: "Nos expertises sont réalisées conformément aux standards d’évaluation publiés par Royal Institution of Chartered Surveyors (RICS).",
                              bold: false,
                              color: "000000",
                              break: 1
                          }),
                          new TextRun({
                              text: "Les méthodologies d’évaluation préconisées pour le bien expertisé sont :",
                      bold: false,
                      color: "000000",
                      break: 1
                  }),

                  new TextRun({
                      text: this.onSubmit('mtd1'),
            bold: false,
            color: "000000",
            break: 1
        }),
        new TextRun({
            text: this.onSubmit('mtd2'),
  bold: false,
  color: "000000",
  break: 1
}),
new TextRun({
    text: this.onSubmit('mtd3'),
bold: false,
color: "000000",
break: 1
}),
                        ],
                          spacing: {
                            after: 200,
                            before : 20
                        }
                        }),

                        new Paragraph({children: [
                            new TextRun({
                                text:"     Hypothéses Particulières  ",
                                bold: true,
                          color: "000000",
                      }),],
                        heading: HeadingLevel.HEADING_1,
                        shading: {type: ShadingType.REVERSE_DIAGONAL_STRIPE,
                            color: "4682B4",
                            fill: "4682B4",},
                        spacing: {after: 200,
                            before : 20}}),

                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: this.onSubmit1('hyp'),
                                  bold: false,
                                  color: "000000",
                              }),],
                              spacing: {
                                after: 200,
                                before : 20
                            }
                            }),
                            new Paragraph({children: [
                                new TextRun({
                                    text:"     III - COMPOSITION  ",
                                    bold: true,
                              color: "000000",
                          }),],
                            heading: HeadingLevel.HEADING_1,
                            shading: {type: ShadingType.REVERSE_DIAGONAL_STRIPE,
                                color: "4682B4",
                                fill: "4682B4",},
                            spacing: {after: 200,
                                before : 20}}),
                                new Paragraph({
                                    children: [
                                        new TextRun({
                                            text: "Nature de l'actif    : ",
                                      bold: true,
                                      color: "000000",
                                  }),
                                  new TextRun({
                                    text: this.nature("ter", "bur", "ime", "com"),
                              bold: false,
                              color: "000000",
                              break: 1
                          }),
                                ],
                                  spacing: {
                                    after: 200,
                                    before : 20
                                }
                                }), 
                                new Paragraph({
                                    children: [
                                        new TextRun({
                                            text: "Surface de l'actif    :",
                                      bold: false,
                                      color: "000000",
                                      break: 1
                                  }),
                                ],
                                  spacing: {
                                    after: 300,
                                    before : 20
                                }
                                }), 
                                new Paragraph({children: [
                                    new TextRun({
                                        text:"     Dénomination     TF     Etage     Surface(m2)",
                                        bold: true,
                                  color: "000000",
                              }),],
                                heading: HeadingLevel.HEADING_1,
                                shading: {type: ShadingType.REVERSE_DIAGONAL_STRIPE,
                                    color: "4682B4",
                                    fill: "4682B4",},
                                spacing: {after: 200,
                                    before : 20}}),
                                    new Paragraph({
                                        children: [
                                            new TextRun({
                                                text: "     " + this.onSubmit('Pptedite') +"     " +this.onSubmit('TF')+"     "+this.onSubmit('surface')+"     "+this.onSubmit('etage'),
                                          bold: false,
                                          color: "000000",
                                          break: 2
                                      }),
                                    ],
                                      spacing: {
                                        after: 300,
                                        before : 20
                                    }
                                    }), 

                                    new Paragraph({
                                        children: [
                                            new TextRun({
                                                text: "Etat des lieux  :",
                                          bold: false,
                                          color: "000000",
                                          break: 1
                                      }),
                                      new TextRun({
                                        text: this.onSubmit('etat'),
                                  bold: false,
                                  color: "000000",
                                  break: 1
                              }),
                              new TextRun({
                                text: "Droit réel ou Charge foncière : ",
                          bold: false,
                          color: "000000",
                          
                      }),
                      new TextRun({
                        text: this.onSubmit('droit'),
                  bold: false,
                  color: "000000",
                  break: 1
              }),
              new TextRun({
                text: "Etat Apparent :",
          bold: false,
          color: "000000",
          break: 1
      }),
      new TextRun({
        text: this.onSubmit('apparent'),
  bold: false,
  color: "000000",
  break: 1
}),
new TextRun({
    text: "Immeuble neuf : ",
bold: false,
color: "000000",
break: 1
}),
new TextRun({
    text: this.onSubmit('neuf'),
bold: false,
color: "000000",
break: 1
}),
new Paragraph({children: [
    new TextRun({
        text:"        IV - ETAT APPARENT dDU BIEN    ",
        bold: true,
  color: "000000",
}),],
heading: HeadingLevel.HEADING_1,
shading: {type: ShadingType.REVERSE_DIAGONAL_STRIPE,
    color: "4682B4",
    fill: "4682B4",},
spacing: {after: 200,
    before : 20}}),],
     }),
     new Paragraph({
        children: [
            new TextRun({
                text:"•	Une visite de l’actif a été effectuée minutieusement afin de vérifier son état actuel et de s’informer de son environnement mitoyen. "+ this.onSubmit('apparent') +" . Un descriptif photographique a été fait en ce sens.",
          bold: false,
          color: "000000",
          break: 2
      }),
    ],
      spacing: {
        after: 300,
        before : 20
    }
    }),


    new Paragraph({
        alignment: AlignmentType.CENTER,
        heading: HeadingLevel.TITLE,
        children: [
            new TextRun({
          text: "CONTEXTE DE L'ETUDE ",
          bold: true,
          italics: true,
          color: '000000',
          allCaps: true,
          size: 50,
          break: 1,
          

            }),
        ],
        spacing: {
            before: 2500,
            after: 9000,
        },
        }),
        new Paragraph({
            children: [
                new TextRun({
                    text: "ans un but de présentation réelle et claire de la situation et de l’évaluation & la valorisation des actifs, nous avons opté pour une démarche méthodologique et distincte afin d’étudier l’ensemble des aspects fondamen- taux de ces derniers.",
              bold: false,
              color: "000000",
              break: 1
          }),
          new TextRun({
              text:"Notre mission consistera en l’évaluation en valeur vénale et marchande des biens selon les normes de la Royal Institution of Chartered Surveyors (RICS), ainsi qu’en l’estimation de la valeur vénale. Notre démarche s’articulera autour des actions suivantes afin d’affiner au mieux l’évaluation menée :",
      bold: false,
      color: "000000",
      break: 1
  }),
  new TextRun({
text: "-    Une étude de l’environnement mitoyen et proche de l’actif ;",
    bold: false,
color: "000000",
break: 1
}),
new TextRun({
            text:"-    Une visite sur site ;",
        bold: false,
    color: "000000",
    break: 1
    }),
    new TextRun({
            text:"-    Une analyse des documents juridiques, urbanistiques et techniques de l’actif ;",
            bold: false,
        color: "000000",
        break: 1
        }),
        new TextRun({
            text: "-	Une étude de marché par segmentation adéquate aux actifs sur l’usage actuel et futur ;",
                bold: false,
            color: "000000",
            break: 1
            }),
            new TextRun({
                 text:"-	Analyse des offres comparables sur le marché et des transactions récentes effectuées dans le périmètre de l’actif ;",
                    bold: false,
                color: "000000",
                break: 1
                }),
            new TextRun({
                 text:"-	Valorisation des actifs en application de la méthode par comparable et par capitalisation de revenu pour la définition de la valeur vénale et marchande ;",
                       bold: false,
                   color: "000000",
                   break: 1
                   }),
        ],
          spacing: {
            after: 300,
            before : 20
        }
        }),
        clientTable,
        table8,
        new Paragraph({
            alignment: AlignmentType.CENTER,
            heading: HeadingLevel.TITLE,
            children: [
                new TextRun({
              text: "PRESENTATION DU BIEN ",
              bold: true,
              italics: true,
              color: '000000',
              allCaps: true,
              size: 50,
              break: 1,
              
    
                }),
            ],
            spacing: {
                before: 2500,
                after: 9000,
            },
            }),
            new Paragraph({
                  
                text:"         ",
                
                heading: HeadingLevel.TITLE,
               
               
                
                shading: {
                    type: ShadingType.REVERSE_DIAGONAL_STRIPE,
                    color: "4682B4",
                    fill: "4682B4",
                },
                spacing: {
                    after: 200
                }
            }),
    new Paragraph({
                text: "\t2-1.	DEFINITION  DU BIEN	 	",
              
                heading: HeadingLevel.HEADING_1,
               
                spacing: {
                    after: 200
                }
            }),
            new Paragraph({
                text: "\t2-2.	VISITE DU BIEN	 	",
              
                heading: HeadingLevel.HEADING_1,
               
                spacing: {
                    after: 200
                }
            }),
            new Paragraph({
                text: "\t2-3.	DESCRIPTION  DU BIEN	 	",
              
                heading: HeadingLevel.HEADING_1,
               
                spacing: {
                    after: 1900
                }
            }),

            new Paragraph({
                children: [
                    new TextRun({
                        text : "Le bien expertisé est "+ this.nature("ter", "bur", "ime", "com")+this.onSubmit('adresse') +", dans immeuble de prestige d’une chronologie récente. La surface exploitable de l’actif est de "+this.onSubmit('surface')+" m2 conformément au certificat de propriété. L’actif expertisé se trouve dans une copropriété, comme la majorité des immeubles qui longent ladite avenue, le bâtiment en question dispose abrite un rez-de-chaussée commercial. Par ailleurs, il est nécessaire de souligner que l’actif est situé au niveau du "+this.onSubmit('etage')+" étage sur un ensemble de type R+.",
                  bold: false,
                  color: "000000",
                  break: 2
              }),
            ],
              spacing: {
                after: 300,
                before : 20
            }
            }),
            new Paragraph({children: [
                new TextRun({
                    text:"        2-1   Définition DU BIEN    ",
                    bold: true,
              color: "000000",
            }),],
            heading: HeadingLevel.HEADING_1,
            shading: {type: ShadingType.REVERSE_DIAGONAL_STRIPE,
                color: "4682B4",
                fill: "4682B4",},
            spacing: {after: 200,
                before : 20}}),
                new Paragraph({children: [
                    new TextRun({
                        text: this.onSubmit1('bon'),
                        bold: false,
                  color: "000000",
                }),],
                spacing: {after: 200,
                    before : 20}
            }),
            new Paragraph({children: [
                new TextRun({
                    text: "Composition du bien : " + this.nature("ter", "bur", "ime", "com"),
                    bold: true,
              color: "000000",
              break: 1
            }),
            new TextRun({
                text: "Surface Totale : " + this.onSubmit('surface'),
                bold: true,
          color: "000000",
        }),
        ],
            spacing: {after: 200,
                before : 20}
        }),
        new Paragraph({children: [
            new TextRun({
                text:"        Atout du Bien    ",
                bold: true,
          color: "000000",
        }),],
        heading: HeadingLevel.HEADING_1,
        shading: {type: ShadingType.REVERSE_DIAGONAL_STRIPE,
            color: "4682B4",
            fill: "4682B4",},
        spacing: {after: 200,
            before : 20}}),
            new Paragraph({children: [
                
            new TextRun({
                text:  this.onSubmit1('atout'),
                bold: true,
          color: "000000",
        }),
        ],
            spacing: {after: 200,
                before : 20}
        }),

        new Paragraph({children: [
            new TextRun({
                text:"        Analyse de l'environement    ",
                bold: true,
          color: "000000",
        }),],
        heading: HeadingLevel.HEADING_1,
        shading: {type: ShadingType.REVERSE_DIAGONAL_STRIPE,
            color: "4682B4",
            fill: "4682B4",},
        spacing: {after: 200,
            before : 20}}),
            new Paragraph({children: [
                
            new TextRun({
                text:  this.onSubmit1('analyse'),
                bold: true,
          color: "000000",
        }),
        ],
            spacing: {after: 200,
                before : 20}
        }),
        new Paragraph({children: [
                
            new TextRun({
                text:  ">	 Description de l’environnement local et du quartier :",
                bold: true,
          color: "000000",
        }),
        ],
            spacing: {after: 200,
                before : 20}
        }),
        situationTable,
        new Paragraph({children: [
                
            new TextRun({
                text:  ">	 Équipements & Services à proximité :",
                bold: true,
          color: "000000",
        }),
        ],
            spacing: {after: 200,
                before : 20}
        }),
                proxTable,
                new Paragraph({children: [
                    new TextRun({
                        text:"        2-2 visite de bien    ",
                        bold: true,
                  color: "000000",
                }),],
                heading: HeadingLevel.HEADING_1,
                shading: {type: ShadingType.REVERSE_DIAGONAL_STRIPE,
                    color: "4682B4",
                    fill: "4682B4",},
                spacing: {after: 200,
                    before : 20}}),
                    new Paragraph({children: [
                
                        new TextRun({
                            text:  ">La visite du site a eu lieu le "+this.onSubmit('date')+" et a permis de contrôler et de confirmer la nature de l’environnement par rapport aux documents urbanistiques qui nous ont été fournis.",
                            bold: false,
                      color: "000000",
                    }),
                    ],
                        spacing: {after: 200,
                            before : 20}
                    }),

                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        heading: HeadingLevel.TITLE,
                        children: [
                            new TextRun({
                          text: "Reportage photographique ",
                          bold: true,
                          italics: true,
                          color: '000000',
                          allCaps: true,
                          size: 50,
                          break: 1,
                          
                
                            }),
                        ],
                        spacing: {
                            before: 2500,
                            after: 9000,
                        },
                        }),
                        new Paragraph({
                            alignment: AlignmentType.CENTER,
                            heading: HeadingLevel.TITLE,
                            children: [
                                new TextRun({
                              text: "EVALUATION DU BIEN ",
                              bold: true,
                              italics: true,
                              color: '000000',
                              allCaps: true,
                              size: 50,
                              break: 1,
                              
                    
                                }),
                            ],
                            spacing: {
                                before: 2500,
                                after: 9000,
                            },
                            }),
                            new Paragraph({
                                text: "\t4-1.	ETAT APPARENT DU BIEN	",
                              
                                heading: HeadingLevel.HEADING_1,
                               
                                spacing: {
                                    after: 200
                                }
                            }),
                            new Paragraph({
                                text: "\t4-2.	SITUATION FONCIERE – JURIDIQUE	 ",
                              
                                heading: HeadingLevel.HEADING_1,
                               
                                spacing: {
                                    after: 200
                                }
                            }),
                            new Paragraph({
                                text: "\t4-3.	ETUDE DE MARCHE	",
                              
                                heading: HeadingLevel.HEADING_1,
                               
                                spacing: {
                                    after: 1900
                                }
                            }),
                            new Paragraph({children: [
                
                                new TextRun({
                                    text: "Dans cette partie de l’étude, il s’avère pertinent de s’attarder sur les points essentiels qui pourront donner une vision plus explicite sur l’actif expertisé. Il s’agit notamment du volet technique juridique, et commercial. Ces trois volets viennent clarifier l’essence même de cette expertise par l’effectivité de l’état apparent du bien et son état des lieux ainsi que l’état du marché mitoyen.Dans le cadre de cette étude et en premier lieu, pour l’étude de la partie technique, nous serons amenés à analyser les différents aspects relatifs à l’état apparent du second œuvre, des finitions et notamment l’état des parties communes de la résidence. Un système de notation sera mis en place dans le but de détermi ner les différentes contraintes. Cette notation sera élaborée suite aux constats généraux que nous aurons observés lors de la visite sur le terrain.",
                                    bold: false,
                              color: "000000",
                            }),
                            ],
                                spacing: {after: 200,
                                    before : 20}
                            }),

                            new Paragraph({children: [
                
                                new TextRun({
                                    text: "     4-1  Etat apparent du bien ",
                                    bold: true,
                              color: "000000",
                            }),
                            ],
                                spacing: {after: 200,
                                    before : 20}
                            }),
                            new Paragraph({children: [
                
                                new TextRun({
                                    text:"L’étude technique tend à identifier l’aspect et les contraintes ayant un impact direct sur l’actif expertisé. Ce volet d’analyse permet de faire un constat tout en examinant l’état apparent et l’état technique du bien sur leurs états apparents sous réserve d’une étude technique approfondie. Un système de notation est adopté en vue d’avoir une vision claire sur l’état du site. Dans ce volet d’analyse, notre objectif est de démontrer si le site bénéficie d’une accessibilité, de branche- ments, de risques de pollutions et autres risques pouvant altérer le site et par conséquent compromettre sa valorisation",
                                    bold: false,
                              color: "000000",
                            }),
                            ],
                                spacing: {after: 200,
                                    before : 20}
                            }),
                            table9,
                            new Paragraph({children: [
                
                                new TextRun({
                                    text: "     4-2.	SITUATION FONCIERE – JURIDIQUE",
                                    bold: true,
                              color: "000000",
                            }),
                            ],
                                spacing: {after: 200,
                                    before : 20}
                            }),
                            new Paragraph({children: [
                
                                new TextRun({
                                    text: "Les informations juridiques consignées sur le certificat de propriété communiqué, se synthétisent de la manière suivante :",
                                    bold: false,
                              color: "000000",
                            }),
                            ],
                                spacing: {after: 200,
                                    before : 20}
                            }),
                            juridiqueTable,
                            new Paragraph({children: [
                
                                new TextRun({
                                    text: "     4-3.	Etude de marché",
                                    bold: true,
                              color: "000000",
                            }),
                            ],
                                spacing: {after: 200,
                                    before : 20}
                            }),
                            new Paragraph({children: [
                
                                new TextRun({
                                    text: "Dans cet axe, notre étude de marché nous a amenée à faire est un travail de collecte des données immobi- lières et d’analyse d’informations dans le but d’identifier les caractéristiques du marché actuel dans l’immo- bilier et notamment à des biens comparables à l’actif expertisé. Le nombre d’offres de bureaux sur le marché de l’Avenue Omar ibnou AL khattab est très reduit. En effet, ce marché propose essentiellement des appartements  de surfaces intermédiaires et subit une grosse demande qui cause la rareté d’offre de bureaux dans le boulevard. Malgré un niveau de prix croissant,  l’avenue Omar Ibn Al Khattab à privilégier lors d’un investissement locatif, notamment si vous souhaitez optimiser la valorisation de votre patrimoine sur le long terme. En effet, l’immobilier locatif détient de nombreux atouts économiques, touristiques et culturels, qui augmentent le prix dss locaux, mais qui font du boulevard une valeur sûre pour les investisseurs vu la rareté des locaux a usage professionél dans le quartier hormis deux ou trois immeuble a usage de bureaux dans un rayon de 500m a 1KM. Notre étude de marché prend en compte plusieurs critères notamment les statistiques en termes de popula- tion dans la zone, le nombre d’offres présentes ainsi que la demande et les durées d’écoulements des actifs mises à la vente. La valeur homogénéisée moyenne dégagée de cette étude de marché a été retenue pour pour les calculs de la présente expertise. L’analyse du marché concurrentiel met en évidence des prix moyens de vente, compris entre 9500 MAD/ m² et 11.500 MAD/m² sur des actifs de type appartement dans la zone.",
                                    bold: true,
                              color: "000000",
                            }),
                            ],
                                spacing: {after: 200,
                                    before : 20}
                            }),
                            table10,



          ],
      }],
  });
  
  // Used to export the file into a .docx file
  // Packer.toBuffer(doc).then((buffer) => {
  //     fs.writeFileSync("My Document.docx", buffer);
  //     console.log("Document created successfully");
  
  // });
  Packer.toBlob(doc).then(blob => {
      console.log(blob);
      saveAs(blob, "Rapport.docx");
      console.log("Document created successfully");
    });
      
    }
}
