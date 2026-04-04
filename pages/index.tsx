import Head from "next/head";
import { useState, type FormEvent } from "react";
import {
  Box,
  Button,
  Container,
  Heading,
  Input,
  List,
  ListItem,
  Stack,
  Table,
  Tbody,
  Td,
  Text,
  Th,
  Thead,
  Tr,
  UnorderedList,
  useToast,
} from "@chakra-ui/react";

type WorkbookAnalysisItem = {
  fileName: string;
  mode: "groq" | "fallback";
  selectedSheets: string[];
  ignoredSheets: string[];
  reason: string;
  warnings: string[];
};

export default function Home() {
  const [files, setFiles] = useState<FileList | null>(null);
  const [pdfName, setPdfName] = useState("");
  const [loading, setLoading] = useState(false);
  const [analysis, setAnalysis] = useState<WorkbookAnalysisItem[]>([]);
  const toast = useToast();

  const handleSubmit = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();

    if (!pdfName.trim()) {
      toast({ title: "Please enter a PDF filename.", status: "warning", duration: 4000, isClosable: true });
      return;
    }

    if (!files || files.length === 0) {
      toast({ title: "Please select at least one Excel file.", status: "warning", duration: 4000, isClosable: true });
      return;
    }

    const formData = new FormData();
    formData.append("pdfName", pdfName.trim());
    Array.from(files).forEach((file) => formData.append("files", file));

    try {
      setLoading(true);
      const response = await fetch("/api/convert", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        const message = await response.text();
        throw new Error(message || "Conversion failed.");
      }

      const rawAnalysis = response.headers.get("X-Workbook-Analysis");
      if (rawAnalysis) {
        try {
          const parsed = JSON.parse(decodeURIComponent(rawAnalysis)) as WorkbookAnalysisItem[];
          setAnalysis(parsed);
        } catch {
          setAnalysis([]);
        }
      } else {
        setAnalysis([]);
      }

      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = `${pdfName.trim() || "combined-excel"}.pdf`;
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.URL.revokeObjectURL(url);
      toast({ title: "PDF ready.", description: "The workbook was analyzed before export. See the summary below.", status: "success", duration: 4000, isClosable: true });
    } catch (error) {
      setAnalysis([]);
      toast({ title: "Error converting files.", description: error instanceof Error ? error.message : "Unknown error.", status: "error", duration: 6000, isClosable: true });
    } finally {
      setLoading(false);
    }
  };

  return (
    <Container maxW="container.md" py={12}>
      <Head>
        <title>Excel PDF - Professional Converter</title>
      </Head>
      <Stack spacing={8} boxShadow="lg" borderRadius="xl" p={8} bg="white">
        <Box>
          <Heading size="2xl" mb={4} bgGradient="linear(to-r, blue.400, purple.500)" bgClip="text">
            Excel PDF
          </Heading>
          <Text fontSize="xl" color="gray.600" fontWeight="medium">
            Upload your Excel workbooks and let our AI organize, sort, and convert them into elegant PDFs.
          </Text>
        </Box>

        <Box as="form" onSubmit={handleSubmit}>
          <Stack spacing={5}>
            <Input
              type="text"
              placeholder="Enter a PDF filename"
              value={pdfName}
              onChange={(e) => setPdfName(e.target.value)}
            />

            <Input
              type="file"
              accept=".xls,.xlsx"
              multiple
              onChange={(e) => setFiles(e.target.files)}
            />

            {files && files.length > 0 && (
              <Box bg="gray.50" p={4} borderRadius="md">
                <Text fontWeight="medium" mb={2}>
                  Selected files:
                </Text>
                <List spacing={2}>
                  {Array.from(files).map((file) => (
                    <ListItem key={file.name}>
                      {file.name} · {(file.size / 1024).toFixed(1)} KB
                    </ListItem>
                  ))}
                </List>
              </Box>
            )}

            <Button type="submit" colorScheme="blue" isLoading={loading} loadingText="Converting">
              Convert to PDF
            </Button>
          </Stack>
        </Box>

        {analysis.length > 0 && (
          <Box bg="gray.50" p={5} borderRadius="lg" borderWidth="1px" borderColor="gray.200">
            <Heading size="md" mb={4}>
              AI Analysis Summary
            </Heading>
            <Table variant="unstyled" border="1px" borderColor="gray.200">
              <Thead>
                <Tr>
                  <Th border="1px" borderColor="gray.200">File Name</Th>
                  <Th border="1px" borderColor="gray.200">Mode</Th>
                  <Th border="1px" borderColor="gray.200">Selected Sheets</Th>
                  <Th border="1px" borderColor="gray.200">Ignored Sheets</Th>
                  <Th border="1px" borderColor="gray.200">Reason</Th>
                  <Th border="1px" borderColor="gray.200">Warnings</Th>
                </Tr>
              </Thead>
              <Tbody>
                {analysis.map((item) => (
                  <Tr key={item.fileName}>
                    <Td>{item.fileName}</Td>
                    <Td>{item.mode === "groq" ? "Groq AI" : "Fallback rules"}</Td>
                    <Td>{item.selectedSheets.join(", ") || "None"}</Td>
                    <Td>{item.ignoredSheets.join(", ") || "None"}</Td>
                    <Td>{item.reason}</Td>
                    <Td>{item.warnings.length > 0 ? item.warnings.join("; ") : "None"}</Td>
                  </Tr>
                ))}
              </Tbody>
            </Table>
          </Box>
        )}
      </Stack>
    </Container>
  );
}
