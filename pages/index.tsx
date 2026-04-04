import Head from "next/head";
import { useState, type FormEvent, useRef } from "react";
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
  Badge,
  useToast,
  Flex,
  Icon,
} from "@chakra-ui/react";
import { motion, AnimatePresence } from "framer-motion";

// Motion components
const MotionBox = motion(Box);
const MotionStack = motion(Stack);

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
  const [isHovering, setIsHovering] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);
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
      
      toast({
        title: "PDF Downloaded",
        description: "Your professional PDF is ready.",
        status: "success",
        duration: 5000,
        isClosable: true,
        variant: "subtle"
      });
    } catch (error) {
      setAnalysis([]);
      toast({
        title: "Error converting files.",
        description: error instanceof Error ? error.message : "Unknown error.",
        status: "error",
        duration: 6000,
        isClosable: true,
      });
    } finally {
      setLoading(false);
    }
  };

  // Background blobs
  const Background = () => (
    <Box position="fixed" top={0} left={0} w="100%" h="100%" zIndex={-1} overflow="hidden" bg="gray.50">
      <MotionBox
        position="absolute"
        top="-10%" left="-10%" w="50vw" h="50vw" borderRadius="full"
        bg="purple.300" filter="blur(120px)" opacity={0.6}
        animate={{ x: [0, 40, 0], y: [0, -30, 0] }}
        transition={{ duration: 15, repeat: Infinity, ease: "easeInOut" }}
      />
      <MotionBox
        position="absolute"
        bottom="-10%" right="-10%" w="50vw" h="50vw" borderRadius="full"
        bg="blue.300" filter="blur(120px)" opacity={0.5}
        animate={{ x: [0, -40, 0], y: [0, 50, 0] }}
        transition={{ duration: 18, repeat: Infinity, ease: "easeInOut" }}
      />
      <MotionBox
        position="absolute"
        top="40%" left="30%" w="30vw" h="30vw" borderRadius="full"
        bg="pink.200" filter="blur(90px)" opacity={0.4}
        animate={{ scale: [1, 1.2, 1], opacity: [0.3, 0.5, 0.3] }}
        transition={{ duration: 10, repeat: Infinity, ease: "easeInOut" }}
      />
    </Box>
  );

  return (
    <>
      <Head>
        <title>Excel PDF - Premium Converter</title>
      </Head>
      <Background />
      
      <Container maxW="container.md" py={16}>
        <MotionStack
          initial={{ opacity: 0, y: 30 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ duration: 0.6, ease: "easeOut" }}
          spacing={8}
        >
          {/* Main Card */}
          <Box
            bg="whiteAlpha.800"
            backdropFilter="blur(20px)"
            boxShadow="0 25px 50px -12px rgba(0, 0, 0, 0.15)"
            borderRadius="3xl"
            p={10}
            border="1px solid"
            borderColor="whiteAlpha.400"
          >
            <Box mb={8} textAlign="center">
              <Heading
                size="2xl"
                mb={4}
                bgGradient="linear(to-r, blue.500, purple.600)"
                bgClip="text"
                fontWeight="900"
                letterSpacing="tight"
              >
                Excel to PDF
              </Heading>
              <Text fontSize="lg" color="gray.600" fontWeight="medium">
                Transform rough spreadsheets into elegant, business-ready documents with smart AI targeting.
              </Text>
            </Box>

            <Box as="form" onSubmit={handleSubmit}>
              <Stack spacing={6}>
                <Box>
                  <Text fontWeight="bold" mb={2} color="gray.700">Document Name</Text>
                  <Input
                    type="text"
                    size="lg"
                    placeholder="e.g. Q3 Financial Report"
                    value={pdfName}
                    onChange={(e) => setPdfName(e.target.value)}
                    bg="white"
                    borderRadius="xl"
                    boxShadow="sm"
                    border="1px"
                    borderColor="gray.200"
                    _focus={{ borderColor: "purple.400", boxShadow: "0 0 0 1px #9f7aea" }}
                  />
                </Box>

                <Box>
                  <Text fontWeight="bold" mb={2} color="gray.700">Source Workbooks</Text>
                  <Box
                    as={motion.div}
                    whileHover={{ scale: 1.01 }}
                    whileTap={{ scale: 0.99 }}
                    onClick={() => fileInputRef.current?.click()}
                    onDragOver={(e) => { e.preventDefault(); setIsHovering(true); }}
                    onDragLeave={() => setIsHovering(false)}
                    onDrop={(e) => {
                      e.preventDefault();
                      setIsHovering(false);
                      if (e.dataTransfer.files?.length) {
                        setFiles(e.dataTransfer.files);
                      }
                    }}
                    p={8}
                    border="2px dashed"
                    borderColor={isHovering ? "purple.400" : "gray.300"}
                    bg={isHovering ? "purple.50" : "whiteAlpha.900"}
                    borderRadius="2xl"
                    textAlign="center"
                    cursor="pointer"
                    transition="all 0.2s"
                  >
                    <Icon viewBox="0 0 24 24" boxSize={10} color={isHovering ? "purple.500" : "gray.400"} mb={2}>
                      <path fill="currentColor" d="M14 2H6C4.89 2 4 2.9 4 4V20C4 21.11 4.89 22 6 22H18C19.11 22 20 21.11 20 20V8L14 2M13 3.5L18.5 9H13V3.5M8 11.03C8 10.46 8.46 10 9.03 10H14.97C15.54 10 16 10.46 16 11.03V15.97C16 16.54 15.54 17 14.97 17H9.03C8.46 17 8 16.54 8 15.97V11.03M9 11V16H15V11H9Z" />
                    </Icon>
                    <Text fontWeight="bold" color="gray.700" fontSize="lg">
                      {files && files.length > 0 ? `${files.length} file(s) selected` : "Drag & Drop or Click"}
                    </Text>
                    <Text color="gray.500" fontSize="sm" mt={1}>
                      Supported formats: .xls, .xlsx
                    </Text>
                  </Box>
                  <Input
                    type="file"
                    accept=".xls,.xlsx"
                    multiple
                    onChange={(e) => setFiles(e.target.files)}
                    display="none"
                    ref={fileInputRef}
                  />
                </Box>

                <AnimatePresence>
                  {files && files.length > 0 && (
                    <MotionBox
                      initial={{ opacity: 0, height: 0 }}
                      animate={{ opacity: 1, height: "auto" }}
                      exit={{ opacity: 0, height: 0 }}
                    >
                      <List spacing={2}>
                        {Array.from(files).map((file, idx) => (
                          <ListItem key={idx} p={3} bg="white" borderRadius="xl" boxShadow="sm" display="flex" alignItems="center">
                            <Badge colorScheme="green" mr={3}>Excel</Badge>
                            <Text flex="1" fontWeight="medium" isTruncated>{file.name}</Text>
                            <Text fontSize="xs" color="gray.500" fontWeight="bold">{(file.size / 1024).toFixed(1)} KB</Text>
                          </ListItem>
                        ))}
                      </List>
                    </MotionBox>
                  )}
                </AnimatePresence>

                <Button
                  type="submit"
                  size="lg"
                  h="60px"
                  borderRadius="xl"
                  bgGradient="linear(to-r, blue.500, purple.600)"
                  color="white"
                  fontWeight="bold"
                  fontSize="lg"
                  _hover={{ bgGradient: "linear(to-r, blue.600, purple.700)", transform: "translateY(-2px)", boxShadow: "xl" }}
                  _active={{ transform: "translateY(0)" }}
                  isLoading={loading}
                  loadingText="Structuring PDF..."
                  w="100%"
                  boxShadow="lg"
                  transition="all 0.2s"
                >
                  Generate Professional PDF
                </Button>
              </Stack>
            </Box>
          </Box>

          {/* Analysis View */}
          <AnimatePresence>
            {analysis && analysis.length > 0 && (
              <MotionBox
                initial={{ opacity: 0, y: 20 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, scale: 0.95 }}
                bg="whiteAlpha.800"
                backdropFilter="blur(20px)"
                borderRadius="3xl"
                p={8}
                boxShadow="0 10px 30px -10px rgba(0,0,0,0.1)"
                border="1px solid"
                borderColor="whiteAlpha.300"
                overflow="hidden"
              >
                <Heading size="md" mb={6} color="gray.800" display="flex" alignItems="center">
                  <Icon viewBox="0 0 24 24" mr={2} color="purple.500"><path fill="currentColor" d="M12 2C6.5 2 2 6.5 2 12C2 17.5 6.5 22 12 22C17.5 22 22 17.5 22 12C22 6.5 17.5 2 12 2M12 20C7.59 20 4 16.41 4 12C4 7.59 7.59 4 12 4C16.41 4 20 7.59 20 12C20 16.41 16.41 20 12 20M11 7H13V13H11V7M11 15H13V17H11V15Z"/></Icon>
                  AI Processing Report
                </Heading>
                <Box overflowX="auto" borderRadius="xl" border="1px" borderColor="gray.100">
                  <Table variant="simple" size="sm" bg="white">
                    <Thead bg="gray.50">
                      <Tr>
                        <Th color="gray.500">File</Th>
                        <Th color="gray.500">Strategy</Th>
                        <Th color="gray.500">Selected</Th>
                        <Th color="gray.500">Ignored</Th>
                        <Th color="gray.500">Log</Th>
                      </Tr>
                    </Thead>
                    <Tbody>
                      {analysis.map((item, idx) => (
                        <Tr key={idx} _hover={{ bg: "gray.50" }}>
                          <Td fontWeight="medium">
                            <Text isTruncated maxW="150px" title={item.fileName}>{item.fileName}</Text>
                          </Td>
                          <Td>
                            <Badge colorScheme={item.mode === "groq" ? "purple" : "orange"} borderRadius="full" px={2}>
                              {item.mode === "groq" ? "AI Vision" : "Fallback"}
                            </Badge>
                          </Td>
                          <Td>
                            <Flex wrap="wrap" gap={1}>
                              {item.selectedSheets.length > 0 ? item.selectedSheets.map(s => (
                                <Badge key={s} colorScheme="green" variant="subtle" borderRadius="md">{s}</Badge>
                              )) : <Text color="gray.400">-</Text>}
                            </Flex>
                          </Td>
                          <Td>
                            <Flex wrap="wrap" gap={1}>
                              {item.ignoredSheets.length > 0 ? item.ignoredSheets.map(s => (
                                <Badge key={s} colorScheme="gray" variant="subtle" borderRadius="md">{s}</Badge>
                              )) : <Text color="gray.400">-</Text>}
                            </Flex>
                          </Td>
                          <Td>
                            <Text fontSize="xs" color="gray.600" title={item.reason} isTruncated maxW="200px">
                              {item.reason}
                            </Text>
                            {item.warnings.map((w, i) => (
                              <Text key={i} fontSize="xs" color="orange.500" title={w} isTruncated maxW="200px">
                                ⚠ {w}
                              </Text>
                            ))}
                          </Td>
                        </Tr>
                      ))}
                    </Tbody>
                  </Table>
                </Box>
              </MotionBox>
            )}
          </AnimatePresence>
        </MotionStack>
      </Container>
    </>
  );
}
