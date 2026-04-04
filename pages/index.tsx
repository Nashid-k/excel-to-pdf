import Head from "next/head";
import { useState, type FormEvent, useRef, useEffect } from "react";
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
  Icon,
  Center,
} from "@chakra-ui/react";
import { motion, AnimatePresence } from "framer-motion";

// Bubbly Motion components
const MotionBox = motion(Box);
const MotionStack = motion(Stack);
const MotionHeading = motion(Heading);
const MotionText = motion(Text);

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
  const [loadingTextIndex, setLoadingTextIndex] = useState(0);
  const [isHovering, setIsHovering] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const toast = useToast();

  const funnyLoadingTexts = [
    "Mathaji is eating your Excel... 🥗",
    "Mathaji is aligning Particulars... ✨",
    "Mathaji is counting all the Money... 💰",
    "Mathaji is consulting the AI spirits... 🔮",
    "Mathaji is crafting your PDF masterpiece... 🎨",
    "Mathaji is wrapping it up for you! 🎁"
  ];

  useEffect(() => {
    let interval: NodeJS.Timeout;
    if (loading) {
      interval = setInterval(() => {
        setLoadingTextIndex((prev) => (prev + 1) % funnyLoadingTexts.length);
      }, 2500);
    } else {
      setLoadingTextIndex(0);
    }
    return () => clearInterval(interval);
  }, [loading]);

  const handleSubmit = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();

    if (!pdfName.trim()) {
      toast({ 
        title: "Oh no! 🙀", 
        description: "Mathaji says you forgot a document name!", 
        status: "warning", 
        duration: 4000, 
        isClosable: true,
        variant: "subtle"
      });
      return;
    }

    if (!files || files.length === 0) {
      toast({ 
        title: "Wait! 🛑", 
        description: "Mathaji needs at least one Excel file to snack on!", 
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
      const response = await fetch("/api/convert", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        const message = await response.text();
        throw new Error(message || "Conversion failed.");
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
        title: "Yay! 🎊",
        description: "Mathaji made your PDF perfectly! Enjoy! 🎈",
        status: "success",
        duration: 5000,
        isClosable: true,
        variant: "solid",
        position: "top"
      });
    } catch (error) {
      toast({
        title: "Oopsie! 🍄",
        description: error instanceof Error ? error.message : "Something went wrong with Mathaji.",
        status: "error",
        duration: 6000,
        isClosable: true,
      });
    } finally {
      setLoading(false);
    }
  };

  // Squishy floating background blobs
  const Background = () => (
    <Box position="fixed" top={0} left={0} w="100%" h="100%" zIndex={-1} overflow="hidden" bg="orange.50">
      <MotionBox
        position="absolute"
        top="-5%" left="-5%" w="60vw" h="60vw" 
        bg="pink.200" filter="blur(100px)" opacity={0.7}
        animate={{ 
          x: [0, 60, 0], 
          y: [0, -40, 0],
          borderRadius: ["40% 60% 70% 30% / 40% 50% 60% 50%", "60% 40% 30% 70% / 50% 60% 40% 60%", "40% 60% 70% 30% / 40% 50% 60% 50%"]
        }}
        transition={{ duration: 20, repeat: Infinity, ease: "easeInOut" }}
      />
      <MotionBox
        position="absolute"
        bottom="-10%" right="-5%" w="50vw" h="50vw" 
        bg="blue.200" filter="blur(110px)" opacity={0.6}
        animate={{ 
          x: [0, -50, 0], 
          y: [0, 40, 0],
          borderRadius: ["60% 40% 30% 70% / 50% 60% 50% 40%", "40% 60% 70% 30% / 60% 40% 60% 50%", "60% 40% 30% 70% / 50% 60% 50% 40%"]
        }}
        transition={{ duration: 25, repeat: Infinity, ease: "easeInOut" }}
      />
      <MotionBox
        position="absolute"
        top="45%" left="35%" w="35vw" h="35vw" 
        bg="yellow.200" filter="blur(90px)" opacity={0.5}
        animate={{ 
          scale: [1, 1.3, 1],
          opacity: [0.3, 0.6, 0.3],
          borderRadius: ["50%", "45% 55% 65% 35% / 55% 45% 55% 45%", "50%"]
        }}
        transition={{ duration: 15, repeat: Infinity, ease: "easeInOut" }}
      />
    </Box>
  );

  return (
    <>
      <Head>
        <title>Mathaji 1.0 - The Cute Converter 🎀</title>
      </Head>
      <Background />
      
      <Container maxW="container.md" py={12}>
        <MotionStack
          initial={{ opacity: 0, scale: 0.8 }}
          animate={{ opacity: 1, scale: 1 }}
          transition={{ 
            type: "spring",
            damping: 12,
            stiffness: 100
          }}
          spacing={8}
        >
          {/* Main Card */}
          <Box
            bg="whiteAlpha.900"
            backdropFilter="blur(15px)"
            boxShadow="0 30px 60px -20px rgba(255, 105, 180, 0.3)"
            borderRadius="50px"
            p={{ base: 6, md: 12 }}
            border="4px solid"
            borderColor="pink.100"
            position="relative"
            overflow="hidden"
          >
            {/* Cute Decorative Sparkles */}
            <Box position="absolute" top="10px" right="20px" fontSize="2xl">✨</Box>
            <Box position="absolute" bottom="20px" left="20px" fontSize="2xl">🌈</Box>

            <Box mb={8} textAlign="center">
              <MotionHeading
                size="3xl"
                mb={4}
                bgGradient="linear(to-r, pink.400, purple.500, blue.400)"
                bgClip="text"
                fontWeight="900"
                letterSpacing="tight"
                animate={{ y: [0, -5, 0] }}
                transition={{ duration: 3, repeat: Infinity, ease: "easeInOut" }}
              >
                Mathaji 1.0
              </MotionHeading>
              <MotionText fontSize="xl" color="gray.600" fontWeight="bold">
                The magical way to turn your messy numbers into beautiful PDFs! 🦄💎
              </MotionText>
            </Box>

            <Box as="form" onSubmit={handleSubmit}>
              <Stack spacing={8}>
                <Box>
                  <Text fontWeight="800" mb={3} color="purple.600" fontSize="lg" ml={2}>
                    What shall we call this gem? 💎
                  </Text>
                  <Input
                    type="text"
                    size="lg"
                    placeholder="e.g. My Super Cool Report 🚀"
                    value={pdfName}
                    onChange={(e) => setPdfName(e.target.value)}
                    bg="purple.50"
                    borderRadius="25px"
                    boxShadow="inner"
                    border="2px solid"
                    borderColor="purple.100"
                    h="65px"
                    _focus={{ borderColor: "pink.400", bg: "white", boxShadow: "0 0 15px pink" }}
                    fontWeight="bold"
                  />
                </Box>

                <Box>
                  <Text fontWeight="800" mb={3} color="blue.600" fontSize="lg" ml={2}>
                    Give Mathaji some Excel snackies! 🍪
                  </Text>
                  <Box
                    as={motion.div}
                    animate={{ 
                      scale: isHovering ? 1.05 : 1,
                      backgroundColor: isHovering ? "#F0F4FF" : "#FFFFFF" 
                    }}
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
                    p={12}
                    border="4px dashed"
                    borderColor={isHovering ? "pink.400" : "blue.100"}
                    borderRadius="40px"
                    textAlign="center"
                    cursor="pointer"
                    transition="all 0.3s"
                    boxShadow={isHovering ? "0 20px 40px -10px rgba(255, 105, 180, 0.4)" : "none"}
                  >
                    <MotionBox
                      animate={{ 
                        rotate: isHovering ? [0, 10, -10, 10, 0] : 0,
                        scale: isHovering ? 1.2 : 1
                      }}
                      transition={{ duration: 0.5 }}
                    >
                      <Center mb={4}>
                        <Box fontSize="5xl">📁</Box>
                      </Center>
                    </MotionBox>
                    <Text fontWeight="900" color="blue.500" fontSize="xl">
                      {files && files.length > 0 ? `OM NOM! ${files.length} FILE(S)!` : "FEED ME EXCEL!"}
                    </Text>
                    <Text color="blue.300" fontSize="md" mt={2} fontWeight="bold">
                      (Drag them here or click!)
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
                      initial={{ opacity: 0, x: -20 }}
                      animate={{ opacity: 1, x: 0 }}
                      exit={{ opacity: 0, x: 20 }}
                    >
                      <List spacing={3}>
                        {Array.from(files).map((file, idx) => (
                          <ListItem 
                            key={idx} 
                            p={4} 
                            bg="pink.50" 
                            borderRadius="20px" 
                            display="flex" 
                            alignItems="center"
                            border="2px solid"
                            borderColor="pink.100"
                          >
                            <Box fontSize="xl" mr={3}>📊</Box>
                            <Text flex="1" fontWeight="800" color="pink.600" isTruncated>{file.name}</Text>
                            <Badge colorScheme="pink" borderRadius="full" px={3}>{(file.size / 1024).toFixed(1)} KB</Badge>
                          </ListItem>
                        ))}
                      </List>
                    </MotionBox>
                  )}
                </AnimatePresence>

                <Button
                  type="submit"
                  size="xl"
                  h="80px"
                  borderRadius="30px"
                  bgGradient="linear(to-r, pink.400, purple.500, blue.400)"
                  color="white"
                  fontWeight="900"
                  fontSize="2xl"
                  boxShadow="0 15px 30px -10px rgba(159, 122, 234, 0.6)"
                  _hover={{ 
                    bgGradient: "linear(to-r, pink.500, purple.600, blue.500)", 
                    transform: "scale(1.05) rotate(1deg)",
                    boxShadow: "0 25px 50px -15px rgba(159, 122, 234, 0.8)" 
                  }}
                  _active={{ transform: "scale(0.95)" }}
                  isLoading={loading}
                  loadingText={funnyLoadingTexts[loadingTextIndex]}
                  w="100%"
                  transition="all 0.3s cubic-bezier(0.175, 0.885, 0.32, 1.275)"
                >
                  MAGIC GENERATE! ✨
                </Button>
              </Stack>
            </Box>
          </Box>

          <Center mt={8}>
            <MotionText
              fontWeight="bold"
              color="gray.400"
              animate={{ opacity: [0.4, 0.8, 0.4] }}
              transition={{ duration: 4, repeat: Infinity }}
            >
              Mathaji is watching over your documents... 🙏✨
            </MotionText>
          </Center>

        </MotionStack>
      </Container>

      <style jsx global>{`
        body {
          cursor: url('https://raw.githubusercontent.com/half-done/cursors/master/cute-pink/pointer.cur'), auto;
        }
        ::selection {
          background: #FFB6C1;
          color: white;
        }
      `}</style>
    </>
  );
}
