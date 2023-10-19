Imports OfficeOpenXml
Imports OfficeOpenXml.Table
Imports System

Namespace EPPlusSamples
    Public Module ToCollectionSample
        Public Sub Run()
            Console.WriteLine("Running sample 2.7 - ToCollection and ToCollectionWithMappings")
            Using package = New ExcelPackage()
                Dim ws = package.Workbook.Worksheets.Add("Persons")
                ' Load the sample data into the worksheet
                Dim range = ws.Cells("A1").LoadFromCollection(Persons, Sub(options)
                                                                           options.PrintHeaders = True
                                                                           options.TableStyle = TableStyles.Dark1
                                                                       End Sub)

                ' ********************************************************
                '  ToCollection. Automaps cell data to class instance     *
                ' ********************************************************

                Console.WriteLine("******* Sample 2.2 - ToCollection ********" & vbLf)

                ' export the data loaded into the worksheet above to a collection
                Dim exportedPersons = range.ToCollection(Of ToCollectionSamplePerson)()

                For Each person In exportedPersons
                    Console.WriteLine("***************************")
                    Console.WriteLine($"Name: {person.FirstName} {person.LastName}")
                    Console.WriteLine($"Height: {person.Height} cm")
                    Console.WriteLine($"Birthdate: {person.BirthDate.ToShortDateString()}")
                Next

                Console.WriteLine()

                ' ********************************************************
                '  ToCollectionWithMappings. Use this method to manually  *
                '  map all or just some of the cells to your class.       *
                ' ********************************************************

                Console.WriteLine("******* Sample 2.2 - ToCollectionWithMappings ********" & vbLf)

                Dim exportedPersons2 = ws.Cells("A1:D4").ToCollectionWithMappings(Function(row)
                                                                                      ' this runs once per row in the range

                                                                                      ' Create an instance of the exported class
                                                                                      Dim person = New ToCollectionSamplePerson()

                                                                                      ' If some of the cells can be automapped, start by automapping the row data to the class
                                                                                      row.Automap(person)

                                                                                      ' Note that you can only use column names as below
                                                                                      ' if options.HeaderRow is set to the 0-based row index
                                                                                      ' of the header row.
                                                                                      person.FirstName = row.GetValue(Of String)("FirstName")

                                                                                      ' get value by the 0-based column index
                                                                                      person.Height = row.GetValue(Of Integer)(2)

                                                                                      ' return the class instance
                                                                                      Return person
                                                                                  End Function, Sub(options) options.HeaderRow = 0)

                For Each person In exportedPersons2
                    Console.WriteLine("***************************")
                    Console.WriteLine($"Name: {person.FirstName} {person.LastName}")
                    Console.WriteLine($"Height: {person.Height} cm")
                    Console.WriteLine($"Birthdate: {person.BirthDate.ToShortDateString()}")
                Next

                Console.WriteLine()

                ' ********************************************************
                '  ToCollection. Using property attributes for mappings,  *
                '  see the ToCollectionSamplePersonAttr class             *
                ' ********************************************************

                ' Load the sample data into a new worksheet
                Dim ws2 = package.Workbook.Worksheets.Add("Ws2")
                Dim range2 = ws2.Cells("A1").LoadFromCollection(PersonsWithAttributes, Sub(options)
                                                                                           options.PrintHeaders = True
                                                                                           options.TableStyle = TableStyles.Dark1
                                                                                       End Sub)

                Console.WriteLine("******* Sample 2.2 - ToCollection using attributes ********" & vbLf)

                ' export the data loaded into the worksheet above to a collection
                Dim exportedPersons3 = range2.ToCollection(Of ToCollectionSamplePersonAttr)()

                For Each person In exportedPersons3
                    Console.WriteLine("***************************")
                    Console.WriteLine($"Name: {person.FirstName} {person.LastName}")
                    Console.WriteLine($"Height: {person.Height} cm")
                    Console.WriteLine($"Birthdate: {person.BirthDate.ToShortDateString()}")
                Next

                Console.WriteLine()

                ' ********************************************************
                '  ToCollection from a table                              *
                ' ********************************************************
                Console.WriteLine("******* Sample 2.2 - ToCollection from a table ********" & vbLf)
                ' Load the sample data a new worksheet
                Dim ws3 = package.Workbook.Worksheets.Add("Ws3")
                Dim tableRange = ws3.Cells("A1").LoadFromCollection(Persons, Sub(options)
                                                                                 options.PrintHeaders = True
                                                                                 options.TableStyle = TableStyles.Dark1
                                                                             End Sub)
                Dim table = ws3.Tables.GetFromRange(tableRange)
                ' export the data loaded into the worksheet above to a collection
                Dim exportedPersons4 = table.ToCollection(Of ToCollectionSamplePerson)()

                For Each person In exportedPersons4
                    Console.WriteLine("***************************")
                    Console.WriteLine($"Name: {person.FirstName} {person.LastName}")
                    Console.WriteLine($"Height: {person.Height} cm")
                    Console.WriteLine($"Birthdate: {person.BirthDate.ToShortDateString()}")
                Next

                Console.WriteLine()

                Console.WriteLine("Sample 2.7 finished.")
                Console.WriteLine()
            End Using
        End Sub
    End Module
End Namespace
