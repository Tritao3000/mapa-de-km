"use client"

import { useState } from "react"
import { format } from "date-fns"
import {
  Download,
  Eye,
  Plus,
  Trash2,
  CalendarIcon,
  Sun,
  Moon,
} from "lucide-react"
import { useTheme } from "next-themes"
import * as XLSX from "xlsx"
import { ScrollArea, ScrollBar } from "@/components/ui/scroll-area"

import { Button } from "@/components/ui/button"
import { Input } from "@/components/ui/input"
import { Label } from "@/components/ui/label"
import {
  Card,
  CardContent,
  CardDescription,
  CardHeader,
  CardTitle,
} from "@/components/ui/card"
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from "@/components/ui/table"
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog"
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select"
import { Calendar } from "@/components/ui/calendar"
import {
  Popover,
  PopoverContent,
  PopoverTrigger,
} from "@/components/ui/popover"

const percentageOptions = [
  { value: 1, label: "100%" },
  { value: 0.75, label: "75%" },
  { value: 0.5, label: "50%" },
  { value: 0.25, label: "25%" },
]

interface Trip {
  day: string
  empresa: string
  dataHoraInicio: string
  dataHoraRegresso: string
  percentage: number
  kms: number
  localidade: string
}

function DateTimePicker({
  value,
  onChange,
  placeholder = "Selecionar data e hora",
}: {
  value: string
  onChange: (dateTime: string) => void
  placeholder?: string
}) {
  const date = value ? new Date(value) : undefined
  const [isOpen, setIsOpen] = useState(false)

  const hours = Array.from({ length: 24 }, (_, i) => i)

  const handleDateSelect = (selectedDate: Date | undefined) => {
    if (selectedDate && date) {
      const newDate = new Date(selectedDate)
      newDate.setHours(date.getHours())
      newDate.setMinutes(date.getMinutes())
      onChange(format(newDate, "yyyy-MM-dd'T'HH:mm"))
    } else if (selectedDate) {
      const newDate = new Date(selectedDate)
      newDate.setHours(9, 0, 0, 0)
      onChange(format(newDate, "yyyy-MM-dd'T'HH:mm"))
    }
  }

  const handleTimeChange = (type: "hour" | "minute", val: string) => {
    if (date) {
      const newDate = new Date(date)
      if (type === "hour") {
        newDate.setHours(parseInt(val))
      } else if (type === "minute") {
        newDate.setMinutes(parseInt(val))
      }
      onChange(format(newDate, "yyyy-MM-dd'T'HH:mm"))
    }
  }

  return (
    <Popover open={isOpen} onOpenChange={setIsOpen}>
      <PopoverTrigger asChild>
        <Button
          variant="outline"
          role="combobox"
          className="w-full justify-start text-left font-normal"
        >
          <CalendarIcon className="mr-2 h-4 w-4" />
          {date ? format(date, "dd/MM/yyyy HH:mm") : <span>{placeholder}</span>}
        </Button>
      </PopoverTrigger>
      <PopoverContent className="w-auto p-0" align="start">
        <div className="sm:flex">
          <Calendar
            mode="single"
            selected={date}
            onSelect={handleDateSelect}
            initialFocus
          />
          <div className="flex flex-col divide-y border-t sm:h-[300px] sm:flex-row sm:divide-x sm:divide-y-0 sm:border-t-0 sm:border-l">
            <ScrollArea className="w-64 sm:h-full sm:w-auto">
              <div className="flex p-2 sm:flex-col">
                {hours.reverse().map((hour) => (
                  <Button
                    key={hour}
                    size="icon"
                    variant={
                      date && date.getHours() === hour ? "default" : "ghost"
                    }
                    className="aspect-square shrink-0 sm:w-full"
                    onClick={() => handleTimeChange("hour", hour.toString())}
                  >
                    {hour.toString().padStart(2, "0")}
                  </Button>
                ))}
              </div>
              <ScrollBar orientation="horizontal" className="sm:hidden" />
            </ScrollArea>
            <ScrollArea className="w-64 sm:h-full sm:w-auto">
              <div className="flex p-2 sm:flex-col">
                {Array.from({ length: 12 }, (_, i) => i * 5).map((minute) => (
                  <Button
                    key={minute}
                    size="icon"
                    variant={
                      date && date.getMinutes() === minute ? "default" : "ghost"
                    }
                    className="aspect-square shrink-0 sm:w-full"
                    onClick={() =>
                      handleTimeChange("minute", minute.toString())
                    }
                  >
                    {minute.toString().padStart(2, "0")}
                  </Button>
                ))}
              </div>
              <ScrollBar orientation="horizontal" className="sm:hidden" />
            </ScrollArea>
          </div>
        </div>
      </PopoverContent>
    </Popover>
  )
}

export default function MapaKmGenerator() {
  const [previewOpen, setPreviewOpen] = useState(false)
  const [nomeFuncionario, setNomeFuncionario] = useState("")
  const [nif, setNif] = useState("")
  const [viatura, setViatura] = useState("")
  const [mes, setMes] = useState("")
  const [trips, setTrips] = useState<Trip[]>([
    {
      day: "",
      empresa: "",
      dataHoraInicio: "",
      dataHoraRegresso: "",
      percentage: 1,
      kms: 0,
      localidade: "",
    },
  ])

  function addTrip() {
    setTrips([
      ...trips,
      {
        day: "",
        empresa: "",
        dataHoraInicio: "",
        dataHoraRegresso: "",
        percentage: 1,
        kms: 0,
        localidade: "",
      },
    ])
  }

  function removeTrip(index: number) {
    if (trips.length > 1) {
      setTrips(trips.filter((_, i) => i !== index))
    }
  }

  function updateTrip(
    index: number,
    field: keyof Trip,
    value: string | number
  ) {
    const newTrips = [...trips]
    newTrips[index] = { ...newTrips[index], [field]: value }
    setTrips(newTrips)
  }

  function generateExcel() {
    const wb = XLSX.utils.book_new()

    const currencyFormat = '"€" #,##0.00'

    const wsData: (string | number | Date | null)[][] = []
    const wsStyles: Record<string, object> = {}
    const cellFormats: { cell: string; value: object }[] = []

    const setStyle = (row: number, col: number, style: object) => {
      const cell = XLSX.utils.encode_cell({ r: row, c: col })
      wsStyles[cell] = style
    }

    wsData.push([
      null,
      ` NOME DO FUNCIONÁRIO:${nomeFuncionario}`,
      null,
      null,
      "CATEGORIA :",
      null,
      null,
      null,
      " ",
      `Mês:${mes}`,
      null,
    ])
    setStyle(0, 1, { bold: true })

    wsData.push([
      null,
      " NIF:",
      nif,
      null,
      "VIATURA:   ",
      viatura,
      null,
      null,
      null,
      " ",
      null,
      null,
    ])

    wsData.push([
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
    ])
    wsData.push([
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
    ])
    wsData.push([
      null,
      "DIA",
      "EMPRESAS",
      "Início",
      null,
      "Regresso",
      null,
      "%",
      "Kms",
      "LOCALIDADE/TRAJECTO",
      null,
    ])
    setStyle(4, 1, { bold: true })
    setStyle(4, 2, { bold: true })
    setStyle(4, 3, { bold: true })
    setStyle(4, 5, { bold: true })
    setStyle(4, 7, { bold: true })
    setStyle(4, 8, { bold: true })
    setStyle(4, 9, { bold: true })

    wsData.push([
      null,
      null,
      "Serviço efectuado c/ direito ajudas de custo",
      "Dia",
      "Hora",
      "Dia",
      "Hora",
      null,
      null,
      null,
      null,
      null,
    ])
    setStyle(5, 2, { bold: true, italic: true })
    setStyle(5, 3, { bold: true })
    setStyle(5, 4, { bold: true })
    setStyle(5, 5, { bold: true })
    setStyle(5, 6, { bold: true })

    const dataStartRow = 6

    trips.forEach((trip, i) => {
      const excelRow = dataStartRow + i + 1
      const dataInicio = trip.dataHoraInicio
        ? new Date(trip.dataHoraInicio)
        : null
      const dataRegresso = trip.dataHoraRegresso
        ? new Date(trip.dataHoraRegresso)
        : null
      const horaInicio = trip.dataHoraInicio
        ? trip.dataHoraInicio.split("T")[1]?.substring(0, 5) || ""
        : ""
      const horaRegresso = trip.dataHoraRegresso
        ? trip.dataHoraRegresso.split("T")[1]?.substring(0, 5) || ""
        : ""

      wsData.push([
        null,
        trip.day || null,
        trip.empresa || null,
        dataInicio,
        horaInicio,
        dataRegresso,
        horaRegresso,
        trip.percentage || null,
        trip.kms || null,
        trip.localidade || null,
        null,
      ])

      if (trip.percentage !== null && trip.percentage !== undefined) {
        cellFormats.push({
          cell: `H${excelRow}`,
          value: { t: "n", v: trip.percentage, z: "0%" },
        })
      }
      if (trip.kms !== null && trip.kms !== undefined) {
        cellFormats.push({
          cell: `I${excelRow}`,
          value: { t: "n", v: trip.kms, z: "#,##0" },
        })
      }
    })

    const lastDataRow = dataStartRow + trips.length + 1

    wsData.push([
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
    ])

    const calcSectionRow = lastDataRow + 2
    wsData.push([
      null,
      null,
      null,
      "Nº Dias/Kms",
      "Valor Diário/Km",
      "Sub-Total",
      null,
      null,
      null,
      null,
      null,
      null,
    ])
    setStyle(calcSectionRow, 3, { bold: true })
    setStyle(calcSectionRow, 4, { bold: true })
    setStyle(calcSectionRow, 5, { bold: true })

    const count100 = trips.filter((t) => t.percentage === 1).length
    const count75 = trips.filter((t) => t.percentage === 0.75).length
    const count50 = trips.filter((t) => t.percentage === 0.5).length
    const count25 = trips.filter((t) => t.percentage === 0.25).length
    const totalKmsCalc = trips.reduce((sum, t) => sum + (t.kms || 0), 0)

    const row1a = calcSectionRow + 2
    wsData.push([
      null,
      "1.a)",
      "   Completas ( 100% )",
      count100,
      0,
      count100 * 0,
      " Extenso : ",
      null,
      null,
      null,
      null,
      " ",
    ])
    setStyle(row1a, 1, { bold: true })
    setStyle(row1a, 3, { align: "center" })
    setStyle(row1a, 5, { numFmt: currencyFormat })

    const rowNac75 = row1a + 1
    wsData.push([
      null,
      "Nac.",
      "   Reduzidas (   75% )",
      count75,
      0,
      count75 * 0,
      null,
      null,
      null,
      null,
      null,
      null,
    ])
    setStyle(rowNac75, 1, { bold: true })
    setStyle(rowNac75, 3, { align: "center" })
    setStyle(rowNac75, 5, { numFmt: currencyFormat })

    const rowNac50 = row1a + 2
    wsData.push([
      null,
      null,
      "   Reduzidas (   50% )",
      count50,
      0,
      count50 * 0,
      null,
      null,
      null,
      null,
      null,
      " ",
    ])
    setStyle(rowNac50, 3, { align: "center" })
    setStyle(rowNac50, 5, { numFmt: currencyFormat })

    const rowNac25 = row1a + 3
    wsData.push([
      null,
      null,
      "   Reduzidas (   25% )",
      count25,
      0,
      count25 * 0,
      "Assinatura:",
      null,
      null,
      null,
      null,
      null,
    ])
    setStyle(rowNac25, 3, { align: "center" })
    setStyle(rowNac25, 5, { numFmt: currencyFormat })

    const row1b = row1a + 5
    wsData.push([
      null,
      "1.b)",
      "   Completas ( 100% )",
      0,
      0,
      0,
      null,
      null,
      null,
      null,
      null,
      null,
    ])
    setStyle(row1b, 1, { bold: true })
    setStyle(row1b, 3, { align: "center" })
    setStyle(row1b, 5, { numFmt: currencyFormat })

    const rowEst75 = row1b + 1
    wsData.push([
      null,
      "Est.",
      "   Reduzidas (   75% )",
      0,
      0,
      0,
      null,
      null,
      null,
      null,
      null,
      null,
    ])
    setStyle(rowEst75, 1, { bold: true })
    setStyle(rowEst75, 3, { align: "center" })
    setStyle(rowEst75, 5, { numFmt: currencyFormat })

    const rowEst50 = row1b + 2
    wsData.push([
      null,
      null,
      "   Reduzidas (   50% )",
      0,
      0,
      0,
      null,
      null,
      null,
      null,
      null,
      null,
    ])
    setStyle(rowEst50, 3, { align: "center" })
    setStyle(rowEst50, 5, { numFmt: currencyFormat })

    const rowEst25 = row1b + 3
    wsData.push([
      null,
      null,
      "   Reduzidas (   25% )",
      0,
      0,
      0,
      null,
      null,
      null,
      null,
      null,
      null,
    ])
    setStyle(rowEst25, 3, { align: "center" })
    setStyle(rowEst25, 5, { numFmt: currencyFormat })

    const emptyRow = row1b + 5
    wsData.push([
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
    ])

    const kmRow = emptyRow + 2
    wsData.push([
      null,
      "2)",
      "km. Percorridos",
      totalKmsCalc,
      0.4,
      totalKmsCalc * 0.4,
      null,
      null,
      null,
      null,
      null,
      null,
    ])
    setStyle(kmRow, 1, { bold: true })
    setStyle(kmRow, 3, { align: "center", numFmt: "#,##0" })
    setStyle(kmRow, 4, { numFmt: currencyFormat })
    setStyle(kmRow, 5, { numFmt: currencyFormat })

    const subRow = kmRow + 2
    wsData.push([
      null,
      "3)",
      "Subsídio de Alimentação",
      0,
      null,
      0,
      null,
      null,
      null,
      null,
      null,
      null,
    ])
    setStyle(subRow, 1, { bold: true })
    setStyle(subRow, 5, { numFmt: currencyFormat })

    const currentDate = new Date().toLocaleDateString("pt-PT", {
      day: "numeric",
      month: "long",
      year: "numeric",
    })
    const totalRow = subRow + 2
    wsData.push([
      null,
      null,
      "TOTAL RECEBIDO ( 1 + 2 - 3 ) ………………………………………………..",
      null,
      null,
      totalKmsCalc * 0.4,
      null,
      null,
      null,
      `Lisboa, ${currentDate}`,
      null,
    ])
    setStyle(totalRow, 2, { bold: true })
    setStyle(totalRow, 5, { bold: true, numFmt: currencyFormat })
    setStyle(totalRow, 9, { italic: true })

    const ws = XLSX.utils.aoa_to_sheet(wsData)

    Object.entries(wsStyles).forEach(([cell, style]) => {
      if (ws[cell]) {
        ws[cell].s = style
      }
    })

    cellFormats.forEach(({ cell, value }) => {
      ws[cell] = value
    })

    ws["!cols"] = [
      { wch: 3 },
      { wch: 8 },
      { wch: 45 },
      { wch: 12 },
      { wch: 8 },
      { wch: 12 },
      { wch: 8 },
      { wch: 6 },
      { wch: 10 },
      { wch: 40 },
      { wch: 3 },
    ]

    ws["!rows"] = []
    for (let i = 0; i < wsData.length; i++) {
      ws["!rows"].push({ hpt: 18 })
    }

    const sheetName = mes
      ? `${mes} ${nomeFuncionario.split(" ")[0] || "Mapa"}`
      : "Mapa Km"
    XLSX.utils.book_append_sheet(wb, ws, sheetName)

    return wb
  }

  function downloadExcel() {
    const wb = generateExcel()
    const fileName = `Mapa_Km_${mes || "Mapa"}_${nomeFuncionario.replace(/\s+/g, "_") || "Funcionário"}.xlsx`
    XLSX.writeFile(wb, fileName)
  }

  const totalKms = trips.reduce((sum, t) => sum + (t.kms || 0), 0)
  const { resolvedTheme, setTheme } = useTheme()

  return (
    <div className="container mx-auto max-w-6xl px-4 py-8">
      <Button
        variant="ghost"
        size="icon"
        className="absolute top-4 right-4"
        onClick={() => setTheme(resolvedTheme === "dark" ? "light" : "dark")}
      >
        {resolvedTheme === "dark" ? (
          <Sun className="h-5 w-5" />
        ) : (
          <Moon className="h-5 w-5" />
        )}
      </Button>
      <div className="mb-8">
        <h1 className="mb-2 text-3xl font-bold">Gerador de Mapa de Km</h1>
        <p className="text-muted-foreground">
          Preencha o formulário para gerar o seu mapa de quilómetros
        </p>
      </div>

      <div className="space-y-8">
        <Card>
          <CardHeader>
            <CardTitle>Dados do Funcionário</CardTitle>
            <CardDescription>Informações pessoais e da viatura</CardDescription>
          </CardHeader>
          <CardContent className="grid grid-cols-1 gap-4 md:grid-cols-2">
            <div>
              <Label htmlFor="nomeFuncionario">Nome do Funcionário</Label>
              <Input
                id="nomeFuncionario"
                placeholder="Nome completo"
                value={nomeFuncionario}
                onChange={(e) => setNomeFuncionario(e.target.value)}
              />
            </div>
            <div>
              <Label htmlFor="nif">NIF</Label>
              <Input
                id="nif"
                placeholder="NIF"
                value={nif}
                onChange={(e) => setNif(e.target.value)}
              />
            </div>
            <div>
              <Label htmlFor="viatura">Viatura</Label>
              <Input
                id="viatura"
                placeholder="Ex: 51-VA-68"
                value={viatura}
                onChange={(e) => setViatura(e.target.value)}
              />
            </div>
            <div>
              <Label htmlFor="mes">Mês</Label>
              <Input
                id="mes"
                placeholder="Ex: Fevereiro 2026"
                value={mes}
                onChange={(e) => setMes(e.target.value)}
              />
            </div>
          </CardContent>
        </Card>

        <Card>
          <CardHeader>
            <CardTitle>Viagens</CardTitle>
            <CardDescription>
              Adicione todas as viagens realizadas durante o mês
            </CardDescription>
          </CardHeader>
          <CardContent className="space-y-4">
            {trips.map((trip, index) => (
              <div key={index} className="rounded-lg border bg-card p-4">
                <div className="mb-4 flex items-center justify-between">
                  <h3 className="font-medium">Viagem #{index + 1}</h3>
                  {trips.length > 1 && (
                    <Button
                      type="button"
                      variant="ghost"
                      size="sm"
                      onClick={() => removeTrip(index)}
                    >
                      <Trash2 className="h-4 w-4" />
                    </Button>
                  )}
                </div>

                <div className="grid grid-cols-1 gap-4 md:grid-cols-2 lg:grid-cols-4">
                  <div>
                    <Label htmlFor={`day-${index}`}>Dia</Label>
                    <Input
                      id={`day-${index}`}
                      type="number"
                      min="1"
                      max="31"
                      placeholder="1-31"
                      value={trip.day}
                      onChange={(e) => updateTrip(index, "day", e.target.value)}
                    />
                  </div>
                  <div className="md:col-span-2 lg:col-span-1">
                    <Label htmlFor={`empresa-${index}`}>
                      Empresa / Serviço
                    </Label>
                    <Input
                      id={`empresa-${index}`}
                      placeholder="Descrição do serviço"
                      value={trip.empresa}
                      onChange={(e) =>
                        updateTrip(index, "empresa", e.target.value)
                      }
                    />
                  </div>
                  <div>
                    <Label htmlFor={`dataHoraInicio-${index}`}>
                      Data e Hora de Início
                    </Label>
                    <DateTimePicker
                      value={trip.dataHoraInicio}
                      onChange={(v) => updateTrip(index, "dataHoraInicio", v)}
                      placeholder="Data e hora de início"
                    />
                  </div>
                  <div>
                    <Label htmlFor={`dataHoraRegresso-${index}`}>
                      Data e Hora de Regresso
                    </Label>
                    <DateTimePicker
                      value={trip.dataHoraRegresso}
                      onChange={(v) => updateTrip(index, "dataHoraRegresso", v)}
                      placeholder="Data e hora de regresso"
                    />
                  </div>
                  <div>
                    <Label htmlFor={`percentage-${index}`}>Percentagem</Label>
                    <Select
                      value={String(trip.percentage)}
                      onValueChange={(v) =>
                        updateTrip(index, "percentage", parseFloat(v))
                      }
                    >
                      <SelectTrigger id={`percentage-${index}`}>
                        <SelectValue />
                      </SelectTrigger>
                      <SelectContent>
                        {percentageOptions.map((opt) => (
                          <SelectItem key={opt.value} value={String(opt.value)}>
                            {opt.label}
                          </SelectItem>
                        ))}
                      </SelectContent>
                    </Select>
                  </div>
                  <div>
                    <Label htmlFor={`kms-${index}`}>Kms</Label>
                    <Input
                      id={`kms-${index}`}
                      type="number"
                      min="0"
                      placeholder="Quilómetros"
                      value={trip.kms || ""}
                      onChange={(e) =>
                        updateTrip(
                          index,
                          "kms",
                          parseFloat(e.target.value) || 0
                        )
                      }
                    />
                  </div>
                  <div className="md:col-span-2 lg:col-span-2">
                    <Label htmlFor={`localidade-${index}`}>
                      Localidade / Trajecto
                    </Label>
                    <Input
                      id={`localidade-${index}`}
                      placeholder="Ex: Lisboa/Porto"
                      value={trip.localidade}
                      onChange={(e) =>
                        updateTrip(index, "localidade", e.target.value)
                      }
                    />
                  </div>
                </div>
              </div>
            ))}

            <Button type="button" variant="outline" onClick={addTrip}>
              <Plus className="h-4 w-4" />
              Adicionar Viagem
            </Button>
          </CardContent>
        </Card>

        <div className="flex justify-end gap-4">
          <Button
            type="button"
            variant="outline"
            onClick={() => setPreviewOpen(true)}
          >
            <Eye className="h-4 w-4" />
            Pré-visualizar
          </Button>
          <Button type="button" onClick={downloadExcel}>
            <Download className="h-4 w-4" />
            Descarregar Excel
          </Button>
        </div>
      </div>

      <Dialog open={previewOpen} onOpenChange={setPreviewOpen}>
        <DialogContent className="max-h-[90vh] min-w-7xl overflow-y-auto">
          <DialogHeader>
            <DialogTitle>Pré-visualização do Mapa de Km</DialogTitle>
            <DialogDescription>
              Confirme os dados antes de descarregar
            </DialogDescription>
          </DialogHeader>

          <div className="space-y-4">
            <Card>
              <CardHeader>
                <CardTitle className="text-lg">Dados do Funcionário</CardTitle>
              </CardHeader>
              <CardContent className="grid grid-cols-2 gap-4">
                <div>
                  <strong>Nome:</strong> {nomeFuncionario}
                </div>
                <div>
                  <strong>NIF:</strong> {nif}
                </div>
                <div>
                  <strong>Viatura:</strong> {viatura}
                </div>
                <div>
                  <strong>Mês:</strong> {mes}
                </div>
              </CardContent>
            </Card>

            <Card>
              <CardHeader>
                <CardTitle className="text-lg">Viagens Registadas</CardTitle>
              </CardHeader>
              <CardContent>
                <Table>
                  <TableHeader>
                    <TableRow>
                      <TableHead>Dia</TableHead>
                      <TableHead>Empresa</TableHead>
                      <TableHead>Início</TableHead>
                      <TableHead>Regresso</TableHead>
                      <TableHead>%</TableHead>
                      <TableHead>Kms</TableHead>
                      <TableHead>Localidade</TableHead>
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    {trips
                      .filter((t) => t.day && t.empresa)
                      .map((trip, i) => (
                        <TableRow key={i}>
                          <TableCell>{trip.day}</TableCell>
                          <TableCell>{trip.empresa}</TableCell>
                          <TableCell>
                            {trip.dataHoraInicio
                              ? format(
                                  new Date(trip.dataHoraInicio),
                                  "dd/MM/yyyy HH:mm"
                                )
                              : ""}
                          </TableCell>
                          <TableCell>
                            {trip.dataHoraRegresso
                              ? format(
                                  new Date(trip.dataHoraRegresso),
                                  "dd/MM/yyyy HH:mm"
                                )
                              : ""}
                          </TableCell>
                          <TableCell>{trip.percentage * 100}%</TableCell>
                          <TableCell>{trip.kms}</TableCell>
                          <TableCell>{trip.localidade}</TableCell>
                        </TableRow>
                      ))}
                    <TableRow>
                      <TableCell colSpan={5} className="text-right font-bold">
                        Total Kms:
                      </TableCell>
                      <TableCell className="font-bold">{totalKms}</TableCell>
                      <TableCell></TableCell>
                    </TableRow>
                  </TableBody>
                </Table>
              </CardContent>
            </Card>

            <Card>
              <CardHeader>
                <CardTitle className="text-lg">Resumo</CardTitle>
              </CardHeader>
              <CardContent className="grid grid-cols-3 gap-4">
                <div>
                  <strong>Total Kms:</strong> {totalKms} km
                </div>
                <div>
                  <strong>Contribuição por Km:</strong> 0,40 €
                </div>
                <div>
                  <strong>Total a Receber:</strong>{" "}
                  {(totalKms * 0.4).toFixed(2)} €
                </div>
              </CardContent>
            </Card>
          </div>

          <div className="mt-4 flex justify-end gap-2">
            <Button variant="outline" onClick={() => setPreviewOpen(false)}>
              Cancelar
            </Button>
            <Button onClick={downloadExcel}>
              <Download className="h-4 w-4" />
              Descarregar Excel
            </Button>
          </div>
        </DialogContent>
      </Dialog>
    </div>
  )
}
