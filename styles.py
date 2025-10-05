from docx import Document

doc = Document("template/blank-template.docx")

# print("Available styles:")
# for s in doc.styles:
#     try:
#         print("-", s.name)
#     except:
#         pass


for s in doc.styles:
    try:
        style_id = s.style_id        # internal ID
        style_name = s.name          # display name
        print(f"ID: {style_id}  |  Name: {style_name}")
    except:
        pass



# ID: Normal  |  Name: Normal
# ID: Heading1  |  Name: Heading 1
# ID: Heading2  |  Name: Heading 2
# ID: Heading3  |  Name: Heading 3
# ID: Heading4  |  Name: Heading 4
# ID: Heading5  |  Name: Heading 5
# ID: Heading6  |  Name: Heading 6
# ID: Heading7  |  Name: Heading 7
# ID: Heading8  |  Name: Heading 8
# ID: Heading9  |  Name: Heading 9
# ID: DefaultParagraphFont  |  Name: Default Paragraph Font
# ID: TableNormal  |  Name: Normal Table
# ID: NoList  |  Name: No List
# ID: AbsTitle  |  Name: AbsTitle*
# ID: Num  |  Name: Num*
# ID: EnRef  |  Name: EnRef*
# ID: Equation  |  Name: Equation*
# ID: FarsiRef  |  Name: FarsiRef*
# ID: Header  |  Name: Header
# ID: HeaderLeft  |  Name: HeaderLeft*
# ID: HeaderRight  |  Name: HeaderRight*
# ID: Footer  |  Name: Footer
# ID: Title24  |  Name: Title 24*
# ID: NormalB  |  Name: NormalB*
# ID: EquaEnd  |  Name: EquaEnd*
# ID: PageNumber  |  Name: page number
# ID: RefB  |  Name: RefB*
# ID: TOC4  |  Name: toc 4
# ID: EquaMid  |  Name: EquaMid*
# ID: EquaStart  |  Name: EquaStart*
# ID: TableTitle  |  Name: Table Title*
# ID: PicTitle  |  Name: Pic Title*
# ID: InPicture  |  Name: In Picture*
# ID: InTable  |  Name: In Table*
# ID: BuletB  |  Name: BuletB*
# ID: NormalWeb  |  Name: Normal (Web)
# ID: TOC1  |  Name: toc 1
# ID: TOC2  |  Name: toc 2
# ID: TOC3  |  Name: toc 3
# ID: Title14  |  Name: Title 14*
# ID: Heading3Char  |  Name: Heading 3 Char
# ID: Heading2Char  |  Name: Heading 2 Char
# ID: FootnoteText  |  Name: footnote text
# ID: Caption  |  Name: Caption
# ID: SubHedList  |  Name: SubHedList*
# ID: NormalLeftB  |  Name: NormalLeftB*
# ID: FootnoteReference  |  Name: footnote reference
# ID: FooterChar  |  Name: Footer Char
# ID: Hyperlink  |  Name: Hyperlink
# ID: NoSpacing  |  Name: No Spacing
# ID: TableofFigures  |  Name: table of figures
# ID: RefItalic  |  Name: RefItalic*
# ID: RefItalicCharChar  |  Name: RefItalic* Char Char
# ID: NormalBCharChar  |  Name: NormalB* Char Char
# ID: TableGrid  |  Name: Table Grid
# ID: Title16  |  Name: Title 16*
# ID: InTableR  |  Name: In Table R*
# ID: Bulet0  |  Name: Bulet*
# ID: Code  |  Name: Code*
# ID: CodeBold  |  Name: CodeBold*
# ID: CodeComment  |  Name: CodeComment*
# ID: CommentReference  |  Name: annotation reference
# ID: TOC5  |  Name: toc 5
# ID: TOC6  |  Name: toc 6
# ID: CommentText  |  Name: annotation text
# ID: Bulet  |  Name: Bulet
# ID: CommentSubject  |  Name: annotation subject
# ID: CodeCharChar  |  Name: Code* Char Char
# ID: CodeBoldCharChar  |  Name: CodeBold* Char Char
# ID: TOC7  |  Name: toc 7
# ID: TOC8  |  Name: toc 8
# ID: TOC9  |  Name: toc 9
# ID: BalloonText  |  Name: Balloon Text
# ID: Title18  |  Name: Title 18*
# ID: NoSpacingChar  |  Name: No Spacing Char
# ID: HeaderChar  |  Name: Header Char
# ID: Style1  |  Name: Style1
# ID: Style2  |  Name: Style2
# ID: Style3  |  Name: Style3
# ID: Style2Char  |  Name: Style2 Char
# ID: Style4  |  Name: Style4
# ID: Style3Char  |  Name: Style3 Char
# ID: ListParagraph  |  Name: List Paragraph
# ID: Style4Char  |  Name: Style4 Char
# ID: Strong  |  Name: Strong
# ID: Emphasis  |  Name: Emphasis
# ID: Title  |  Name: Title
# ID: TitleChar  |  Name: Title Char
# ID: StyleCentered  |  Name: Style Centered
# ID: Style3Char  |  Name: Style3 Char
# ID: ListParagraph  |  Name: List Paragraph
# ID: Style4Char  |  Name: Style4 Char
# ID: Strong  |  Name: Strong
# ID: Emphasis  |  Name: Emphasis
# ID: Title  |  Name: Title
# ID: TitleChar  |  Name: Title Char
# ID: StyleCentered  |  Name: Style Centered
# ID: Style4Char  |  Name: Style4 Char
# ID: Strong  |  Name: Strong
# ID: Emphasis  |  Name: Emphasis
# ID: Title  |  Name: Title
# ID: TitleChar  |  Name: Title Char
# ID: StyleCentered  |  Name: Style Centered
# ID: TitleChar  |  Name: Title Char
# ID: StyleCentered  |  Name: Style Centered