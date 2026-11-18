import Head from "next/head";
import { useState, type FormEvent, useRef, useEffect, type DragEvent } from "react";
import {
  Box,
  Button,
  Container,
  Heading,
  Input,
  List,
  ListItem,
  Stack,
  Text,
  Badge,
  useToast,
  Center,
  useColorMode,
  useColorModeValue,
  IconButton,
} from "@chakra-ui/react";
import { motion, AnimatePresence } from "framer-motion";

// Clean Motion components
const MotionBox = motion(Box);
const MotionStack = motion(Stack);
const MotionHeading = motion(Heading);
const MotionText = motion(Text);

// Animated Theme Switcher Component
const ThemeSwitcher = () => {
  const { colorMode, toggleColorMode } = useColorMode();
  const iconColor = useColorModeValue("orange.400", "yellow.300");
  const bg = useColorModeValue("white", "whiteAlpha.100");
  const border = useColorModeValue("#D2D2D7", "#3A3A3C");

  return (
    <MotionBox
      initial={{ opacity: 0, y: -20 }}
      animate={{ opacity: 1, y: 0 }}
      transition={{ delay: 0.5 }}
      position="absolute"
      top={{ base: "20px", md: "40px" }}
      right={{ base: "20px", md: "40px" }}
      zIndex={100}
    >
      <IconButton
        aria-label="Toggle Theme"
        icon={
          <MotionBox
            key={colorMode}
            initial={{ rotate: -90, opacity: 0, scale: 0.5 }}
            animate={{ rotate: 0, opacity: 1, scale: 1 }}
            transition={{ type: "spring", damping: 12 }}
            fontSize="xl"
          >
            {colorMode === "light" ? "🌙" : "☀️"}
          </MotionBox>
        }
        onClick={toggleColorMode}
        variant="outline"
        borderRadius="14px"
        bg={bg}
        borderColor={border}
        _hover={{ bg: useColorModeValue("gray.50", "whiteAlpha.200"), transform: "scale(1.05)" }}
        _active={{ transform: "scale(0.95)" }}
        size="lg"
        boxShadow="sm"
      />
    </MotionBox>
  );
};

export default function Home() {
  const [files, setFiles] = useState<FileList | null>(null);
  const [pdfName, setPdfName] = useState("");
  const [loading, setLoading] = useState(false);
  const [loadingTextIndex, setLoadingTextIndex] = useState(0);
  const [isHovering, setIsHovering] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const toast = useToast();
  const { colorMode } = useColorMode();

  // Responsive Colors
  const bgColor = useColorModeValue("#F5F5F7", "#000000");
  const cardBg = useColorModeValue("white", "#1C1C1E");
  const cardBorder = useColorModeValue("#D2D2D7", "#3A3A3C");
  const headingColor = useColorModeValue("#1D1D1F", "#F5F5F7");
  const subTextColor = useColorModeValue("#86868B", "#A1A1A6");
  const inputBg = useColorModeValue("#F5F5F7", "#2C2C2E");
  const dropzoneBg = useColorModeValue("#F5F5F7", "#1C1C1E");
  const loaderOverlayBg = useColorModeValue("rgba(255, 255, 255, 0.9)", "rgba(28, 28, 30, 0.9)");

  const funnyLoadingTexts = [
    "Mathaji Excel thinnukayaanu... 🥙",
    "Mathaji is aligning Particulars... ✨",
    "PDF varum, onnu wait cheyyu... ⏳",
    "Kanakku kootti thalapennu! 🤯",
    "Polichadukkaan ready aayikkolu... 🚀",
    "Mathaji is consulting the AI spirits... 🔮",
    "Shsh.. Mathaji dhyanathilaanu... 🧘",
    "hampir ready aayi! 🍫",
    "Mathaji crafting your PDF masterpiece... 🎨",
    "Basking in your document's aura... ✨"
  ];

  useEffect(() => {
    let interval: NodeJS.Timeout;
    if (loading) {
      interval = setInterval(() => {
        setLoadingTextIndex((prev) => (prev + 1) % funnyLoadingTexts.length);
      }, 1500);
    } else {
      setLoadingTextIndex(0);
    }
    return () => clearInterval(interval);
  }, [loading]);

  const handleSubmit = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();

    if (!pdfName.trim()) {
      toast({ 
        title: "Missing Name", 
        description: "Please provide a name for your masterpiece.", 
        status: "warning", 
        duration: 4000, 
        isClosable: true,
        variant: "subtle"
      });
      return;
    }

    if (!files || files.length === 0) {
      toast({ 
        title: "No Files", 
        description: "Mathaji needs at least one Excel file to process.", 
        status: "warning", 
        duration: 4000, 
        isClosable: true,
        variant: "subtle"
      });
      return;
    }

    const formData = new FormData();
    formData.append("pdfName", pdfName.trim());
    Array.from(files).forEach((file) => formData.append("files", file));

    try {
      setLoading(true);
      const startTime = Date.now();
      
      const response = await fetch("/api/convert", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        const message = await response.text();
        throw new Error(message || "Conversion failed.");
      }

      const blob = await response.blob();
      
      const elapsed = Date.now() - startTime;
      if (elapsed < 5000) {
        await new Promise((resolve) => setTimeout(resolve, 5000 - elapsed));
      }

      const url = window.URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = `${pdfName.trim() || "combined-excel"}.pdf`;
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.URL.revokeObjectURL(url);
      
      toast({
        title: "Done! 🎊",
        description: "Your PDF is ready for the world.",
        status: "success",
        duration: 5000,
        isClosable: true,
        variant: "solid",
        position: "top"
      });
    } catch (error) {
      toast({
        title: "Error",
        description: error instanceof Error ? error.message : "Something went wrong.",
        status: "error",
        duration: 6000,
        isClosable: true,
      });
    } finally {
      setLoading(false);
    }
  };

  return (
    <>
      <Head>
        <title>MATHAJI 1.0 - Professional PDF Engine</title>
      </Head>
      <Box position="fixed" top={0} left={0} w="100%" h="100%" zIndex={-1} bg={bgColor} transition="background-color 0.4s ease" />
      
      <ThemeSwitcher />

      <Container maxW={{ base: "95%", sm: "container.sm", md: "container.md" }} py={{ base: 8, md: 24 }}>
        <MotionStack
          initial={{ opacity: 0, scale: 0.99 }}
          animate={{ opacity: 1, scale: 1 }}
          transition={{ duration: 0.8, ease: [0.22, 1, 0.36, 1] }}
          spacing={{ base: 8, md: 16 }}
        >
          {/* Main Card */}
          <Box
            bg={cardBg}
            boxShadow={colorMode === "light" ? "0 8px 40px rgba(0,0,0,0.04)" : "none"}
            borderRadius="24px"
            border="1px solid"
            borderColor={cardBorder}
            p={{ base: 6, md: 16 }}
            position="relative"
            overflow="hidden"
            transition="all 0.4s ease"
          >
            {/* Minimalist Header */}
            <Box mb={{ base: 8, md: 16 }} textAlign="center">
              <Text 
                fontSize="xs" 
                fontWeight="700" 
                color="#0066CC" 
                letterSpacing="0.2em" 
                mb={2}
                textTransform="uppercase"
              >
                DOCUMENT ENGINE
              </Text>
              <MotionHeading
                size={{ base: "2xl", md: "3xl" }}
                mb={4}
                color={headingColor}
                fontWeight="700"
                letterSpacing="-0.03em"
              >
                MATHAJI 1.0
              </MotionHeading>
              <Text fontSize={{ base: "lg", md: "xl" }} color={subTextColor} fontWeight="500">
                നിങ്ങളുടെ കണക്കുകൾ ശരിയാക്കാം! 🦄💎
              </Text>
            </Box>

            <Box as="form" onSubmit={handleSubmit} position="relative">
              <Stack spacing={{ base: 6, md: 10 }}>
                <Box>
                  <Text fontWeight="600" mb={3} color={headingColor} fontSize="sm" ml={1}>
                    REPORT NAME
                  </Text>
                  <Input
                    type="text"
                    h={{ base: "52px", md: "64px" }}
                    bg={inputBg}
                    border="1px solid"
                    borderColor={cardBorder}
                    borderRadius="14px"
                    placeholder="e.g. Monthly Vouchers"
                    value={pdfName}
                    onChange={(e) => setPdfName(e.target.value)}
                    _focus={{ borderColor: "#0066CC", bg: colorMode === "light" ? "white" : "#3A3A3C", boxShadow: "0 0 0 4px rgba(0, 102, 204, 0.1)" }}
                    fontSize="md"
                    fontWeight="500"
                    color={headingColor}
                  />
                </Box>

                <Box>
                  <Text fontWeight="600" mb={3} color={headingColor} fontSize="sm" ml={1}>
                    SOURCE FILES
                  </Text>
                  <MotionBox
                    layout
                    whileHover={{ scale: 1.005 }}
                    onClick={() => fileInputRef.current?.click()}
                    onDragOver={(e: DragEvent<HTMLDivElement>) => { e.preventDefault(); setIsHovering(true); }}
                    onDragLeave={() => setIsHovering(false)}
                    onDrop={(e: DragEvent<HTMLDivElement>) => {
                      e.preventDefault();
                      setIsHovering(false);
                      if (e.dataTransfer.files?.length) {
                        setFiles(e.dataTransfer.files);
                      }
                    }}
                    p={{ base: 8, md: 14 }}
                    border="1px dashed"
                    borderColor={isHovering ? "#0066CC" : cardBorder}
                    bg={isHovering ? (colorMode === "light" ? "rgba(0, 102, 204, 0.02)" : "whiteAlpha.50") : dropzoneBg}
                    borderRadius="20px"
                    textAlign="center"
                    cursor="pointer"
                    style={{ transition: "all 0.2s" }}
                  >
                    <Box fontSize="4xl" mb={4} opacity={0.6}>📊</Box>
                    <Text fontWeight="600" color={headingColor} fontSize="lg">
                      {files && files.length > 0 ? `${files.length} ഫയലുകൾ ചേർത്തു` : "Excel ഫയലുകൾ തിരഞ്ഞെടുക്കുക"}
                    </Text>
                    <Text color={subTextColor} fontSize="sm" mt={2} fontWeight="500">
                      Drag & Drop or Click to Browse
                    </Text>
                  </MotionBox>
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
                      initial={{ opacity: 0, y: 5 }}
                      animate={{ opacity: 1, y: 0 }}
                      exit={{ opacity: 0, y: -5 }}
                    >
                      <List spacing={2}>
                        {Array.from(files).map((file, idx) => (
                          <ListItem 
                            key={idx} 
                            p={3} 
                            bg={colorMode === "light" ? "white" : "whiteAlpha.50"} 
                            borderRadius="12px" 
                            display="flex" 
                            alignItems="center"
                            border="1px solid"
                            borderColor={cardBorder}
                          >
                            <Box fontSize="lg" mr={3}>📄</Box>
                            <Text flex="1" fontWeight="600" color={headingColor} fontSize="sm" isTruncated>{file.name}</Text>
                            <Badge variant="subtle" colorScheme="gray" borderRadius="6px" fontSize="xs">
                              {(file.size / 1024).toFixed(1)} KB
                            </Badge>
                          </ListItem>
                        ))}
                      </List>
                    </MotionBox>
                  )}
                </AnimatePresence>

                <Button
                  type="submit"
                  h={{ base: "60px", md: "72px" }}
                  borderRadius="16px"
                  bg="#0066CC"
                  color="white"
                  fontWeight="600"
                  fontSize="lg"
                  _hover={{ bg: "#0071E3", transform: "translateY(-1px)" }}
                  _active={{ bg: "#0055AA", transform: "translateY(0)" }}
                  isLoading={loading}
                  isDisabled={loading}
                  w="100%"
                  transition="all 0.2s"
                >
                  {loading ? "Processing..." : "Generate PDF"}
                </Button>
              </Stack>

              {/* Minimal Loading Overlay */}
              <AnimatePresence>
                {loading && (
                  <MotionBox
                    initial={{ opacity: 0 }}
                    animate={{ opacity: 1 }}
                    exit={{ opacity: 0 }}
                    position="absolute"
                    top="-16px"
                    left="-16px"
                    right="-16px"
                    bottom="-16px"
                    bg={loaderOverlayBg}
                    borderRadius="24px"
                    zIndex={10}
                    display="flex"
                    flexDirection="column"
                    alignItems="center"
                    justifyContent="center"
                    backdropFilter="blur(5px)"
                  >
                    <MotionBox
                      animate={{ rotate: 360 }}
                      transition={{ duration: 2, repeat: Infinity, ease: "linear" }}
                      w="40px"
                      h="40px"
                      border="3px solid"
                      borderColor={cardBorder}
                      borderTopColor="#0066CC"
                      borderRadius="full"
                      mb={6}
                    />
                    <AnimatePresence mode="wait">
                      <MotionBox
                        key={loadingTextIndex}
                        initial={{ opacity: 0, y: 10 }}
                        animate={{ opacity: 1, y: 0 }}
                        exit={{ opacity: 0, y: -10 }}
                        transition={{ duration: 0.4 }}
                        textAlign="center"
                        px={8}
                      >
                        <Text fontSize="xl" fontWeight="700" color={headingColor} mb={1}>
                          {funnyLoadingTexts[loadingTextIndex]}
                        </Text>
                        <Text fontSize="sm" color={subTextColor} fontWeight="600">
                          Please keep this window open
                        </Text>
                      </MotionBox>
                    </AnimatePresence>
                  </MotionBox>
                )}
              </AnimatePresence>
            </Box>
          </Box>

          <Center>
            <Text fontSize="xs" fontWeight="600" color={subTextColor} letterSpacing="0.05em">
              BUILT FOR PRECISION & SPEED
            </Text>
          </Center>
        </MotionStack>
      </Container>

      <style jsx global>{`
        body {
          -webkit-font-smoothing: antialiased;
          -moz-osx-font-smoothing: grayscale;
          transition: background-color 0.4s ease;
        }
      `}</style>
    </>
  );
}
