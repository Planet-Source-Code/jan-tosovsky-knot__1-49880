Attribute VB_Name = "knot"
Option Explicit
Type glCoord
    X As GLfloat
    Y As GLfloat
    Z As GLfloat
End Type

Type glColors
    R As GLfloat
    G As GLfloat
    B As GLfloat
End Type

Type TmyObject
    numfaces As Long
    numverts As Long
    numsides As Long
    faces() As Long
    vertices() As glCoord
    colors() As glColors
    B As GLfloat
End Type

Public Const KnotDL = 1

Sub createKnot(SCALING_FACTOR1 As Long, SCALING_FACTOR2 As Long, RADIUS1 As Single, RADIUS2 As Single, RADIUS3 As Single)
    Dim Knot As TmyObject, pi As Double
    Dim I As Long, J As Long, Count1 As Long, Count2 As Long
    Dim Alpha As GLfloat, Beta As GLfloat, Rotation As GLfloat
    Dim Distance As GLfloat, MinDistance As GLfloat
    Dim X As GLfloat, Y As GLfloat, Z As GLfloat
    Dim dx As GLfloat, dy As GLfloat, dz As GLfloat
    Dim Value  As GLfloat, modulus As GLfloat, dist As GLfloat
    Dim index1 As Long, index2 As Long

    pi = Atn(1) * 4
    Knot.numsides = 4
    
    ReDim Knot.vertices(SCALING_FACTOR1 * SCALING_FACTOR2)
    ReDim Knot.colors(SCALING_FACTOR1 * SCALING_FACTOR2)
    ReDim Knot.faces(SCALING_FACTOR1 * SCALING_FACTOR2 * 4 * 4)
    
    Alpha = 0
    For Count2 = 0 To SCALING_FACTOR2 - 1
        Alpha = Alpha + 2 * pi / SCALING_FACTOR2
        X = RADIUS2 * Cos(2 * Alpha) + RADIUS1 * Sin(Alpha)
        Y = RADIUS2 * Sin(2 * Alpha) + RADIUS1 * Cos(Alpha)
        Z = RADIUS2 * Cos(3 * Alpha)
        dx = -2 * RADIUS2 * Sin(2 * Alpha) + RADIUS1 * Cos(Alpha)
        dy = 2 * RADIUS2 * Cos(2 * Alpha) - RADIUS1 * Sin(Alpha)
        dz = -3 * RADIUS2 * Sin(3 * Alpha)
        Value = Sqr(dx * dx + dz * dz)
        modulus = Sqr(dx * dx + dy * dy + dz * dz)
        
        Beta = 0
        For Count1 = 0 To SCALING_FACTOR1 - 1
            
            Beta = Beta + 2 * pi / SCALING_FACTOR1
            
            Knot.vertices(Knot.numverts).X = X - RADIUS3 * (Cos(Beta) * dz - Sin(Beta) * dx * dy / modulus) / Value
            Knot.vertices(Knot.numverts).Y = Y - RADIUS3 * Sin(Beta) * Value / modulus
            Knot.vertices(Knot.numverts).Z = Z + RADIUS3 * (Cos(Beta) * dx + Sin(Beta) * dy * dz / modulus) / Value
    
            dist = Sqr(Knot.vertices(Knot.numverts).X * Knot.vertices(Knot.numverts).X + _
                       Knot.vertices(Knot.numverts).Y * Knot.vertices(Knot.numverts).Y + _
                       Knot.vertices(Knot.numverts).Z * Knot.vertices(Knot.numverts).Z)
            
            Knot.colors(Knot.numverts).R = ((2 / dist) + (0.5 * Sin(Beta) + 0.4)) / 2#
            Knot.colors(Knot.numverts).G = ((2 / dist) + (0.5 * Sin(Beta) + 0.4)) / 2#
            Knot.colors(Knot.numverts).B = ((2 / dist) + (0.5 * Sin(Beta) + 0.4)) / 2#
            
            Knot.numverts = Knot.numverts + 1
        Next Count1
    Next Count2
  
    For Count1 = 0 To SCALING_FACTOR2 - 1
        index1 = Count1 * SCALING_FACTOR1
        index2 = index1 + SCALING_FACTOR1
        index2 = index2 Mod Knot.numverts
        Rotation = 0
        MinDistance = (Knot.vertices(index1).X - Knot.vertices(index2).X) * (Knot.vertices(index1).X - Knot.vertices(index2).X) + _
                      (Knot.vertices(index1).Y - Knot.vertices(index2).Y) * (Knot.vertices(index1).Y - Knot.vertices(index2).Y) + _
                      (Knot.vertices(index1).Z - Knot.vertices(index2).Z) * (Knot.vertices(index1).Z - Knot.vertices(index2).Z)
    
        For Count2 = 1 To SCALING_FACTOR1 - 1
            index2 = Count2 + index1 + SCALING_FACTOR1
            If Count1 = SCALING_FACTOR2 - 1 Then index2 = Count2
            Distance = (Knot.vertices(index1).X - Knot.vertices(index2).X) * (Knot.vertices(index1).X - Knot.vertices(index2).X) + _
                       (Knot.vertices(index1).Y - Knot.vertices(index2).Y) * (Knot.vertices(index1).Y - Knot.vertices(index2).Y) + _
                       (Knot.vertices(index1).Z - Knot.vertices(index2).Z) * (Knot.vertices(index1).Z - Knot.vertices(index2).Z)
            If Distance < MinDistance Then
                MinDistance = Distance
                Rotation = Count2
            End If
        Next Count2

        For Count2 = 0 To SCALING_FACTOR1 - 1
            Knot.faces(4 * (index1 + Count2) + 0) = index1 + Count2
            
            index2 = Count2 + 1
            index2 = index2 Mod SCALING_FACTOR1
            Knot.faces(4 * (index1 + Count2) + 1) = index1 + index2
            
            index2 = Round(Count2 + Rotation + 1)
            index2 = index2 Mod SCALING_FACTOR1
            Knot.faces(4 * (index1 + Count2) + 2) = (index1 + index2 + SCALING_FACTOR1) Mod Knot.numverts
            
            index2 = Round(Count2 + Rotation)
            index2 = index2 Mod SCALING_FACTOR1
            
            Knot.faces(4 * (index1 + Count2) + 3) = (index1 + index2 + SCALING_FACTOR1) Mod Knot.numverts
            Knot.numfaces = Knot.numfaces + 1
        Next Count2
    Next Count1

    glNewList KnotDL, GL_COMPILE
        glBegin (GL_QUADS)
            J = Knot.numfaces * 4
            For I = 0 To J
                Count1 = Knot.faces(I)
                glColor3f Knot.colors(Count1).R * 2, Knot.colors(Count1).G, Knot.colors(Count1).B
                glVertex3f Knot.vertices(Count1).X, Knot.vertices(Count1).Y, Knot.vertices(Count1).Z
            Next I
        glEnd
    glEndList
End Sub
